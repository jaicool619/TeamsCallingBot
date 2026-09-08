namespace TeamsCallingBot.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Timers;
    using Microsoft.Graph;
    using Microsoft.Graph.Communications.Calls;
    using Microsoft.Graph.Communications.Calls.Media;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Microsoft.Graph.Communications.Resources;
    using Microsoft.Skype.Bots.Media;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using TeamsCallingBot.Audio;
    using TeamsCallingBot.Chat;
    using TeamsCallingBot.Common;
    using TeamsCallingBot.Config;
    using TeamsCallingBot.Mom;
    using TeamsCallingBot.Storage;
    using TeamsCallingBot.Tda;
    using TeamsCallingBot.Video;

    /// <summary>
    /// Complete call lifecycle handler supporting:
    /// 1. Bidirectional Audio (Listen + Speak via AudioSender with real-time speech responses)
    /// 2. Mute / Unmute
    /// 3. Screen share (VBSS) + participant camera VIDEO RECORDING to MJPEG AVI files + snapshots
    /// 4. Disk Storage (Audio, Video, Snapshots, Transcripts, Timeline, MoM Word Doc, Metadata) in one session folder
    /// 5. Bot Video Broadcast (SafeVideoMediaBuffer streaming status card into the meeting)
    /// 6. Auto Leave (Leaves after 30s when all human participants leave)
    /// 7. Chat notifications via Bot Framework Connector (fallback: Graph chat API)
    /// 8. Minutes of Meeting (MoM) Word document generation with attendees, action items, and screenshots
    /// </summary>
    public class CallHandler : HeartbeatHandler
    {
        private int endHandled;
        private readonly IGraphLogger graphLogger;
        private readonly DateTime sessionStartTime;
        private readonly BotOptions options;

        // Media Sockets
        private readonly IAudioSocket audioSocket;
        private readonly IVideoSocket vbssSocket;
        private readonly IVideoSocket videoSocket;

        // Controllers, Timeline & Storage
        public AudioSender AudioSender { get; }
        public AudioAggregator AudioAggregator { get; }
        public RecordingsManager RecordingsManager { get; }
        public MeetingTimeline Timeline { get; }

        // State Flags
        public bool IsMuted { get; private set; } = true;
        private volatile bool isVideoSendActive = false;
        private int joinWelcomeSent = 0;
        private volatile bool callEstablished = false;

        // Fallback photo throttling (only used when video recording is disabled)
        private DateTime lastPhotoTime = DateTime.MinValue;
        private readonly TimeSpan photoInterval = TimeSpan.FromSeconds(5);
        private readonly object photoLock = new object();
        private Bitmap latestScreenBitmap = null;

        // Video recording state (screen share + camera)
        private readonly object videoLock = new object();
        private VideoRecorder vbssRecorder;
        private MediaSessionRecord vbssSession;
        private uint currentVbssMsi;
        private VideoRecorder cameraRecorder;
        private MediaSessionRecord cameraSession;
        private uint currentCameraMsi;
        private int vbssMsiMismatchLogged;
        private readonly HashSet<string> participantsWithUpdateHook = new HashSet<string>();

        // Bot Video Streaming
        private CancellationTokenSource videoBroadcastCts;
        private volatile VideoFormat preferredSendFormat = VideoFormat.NV12_1280x720_15Fps;
        private volatile string activityLine;
        private int broadcastErrorsLogged;

        // Periodic Audio Flush
        private CancellationTokenSource periodicFlushCts;

        // Auto-Leave Timer
        private CancellationTokenSource autoLeaveCts;
        private const int AutoLeaveDebounceSeconds = 30;
        private const int InitialGracePeriodSeconds = 45;

        // Chat Info & Clients
        private readonly string chatThreadId;
        private readonly string botAccessToken;
        private readonly BotFrameworkChatClient chatClient;
        private CancellationTokenSource chatMonitorCts;

        // Real-Time Voice Speech Recognition
        private readonly RealtimeSpeechRecognizer voiceRecognizer;

        // TDA & Visualization
        private readonly TdaClient tdaClient;
        private readonly object visualizationLock = new object();
        private Bitmap currentVisualization;
        private string currentVisualizationTitle;
        private DateTime visualizationExpiresAt = DateTime.MinValue;

        public ICall Call { get; }

        public CallHandler(ICall call, IGraphLogger logger, string chatThreadId = null, string accessToken = null)
            : base(TimeSpan.FromMinutes(1), logger)
        {
            this.Call = call ?? throw new ArgumentNullException(nameof(call));
            this.graphLogger = logger;
            this.sessionStartTime = DateTime.Now;
            this.chatThreadId = chatThreadId;
            this.options = BotOptions.Current ?? new BotOptions();
            this.botAccessToken = accessToken ?? this.options.OverrideBearerToken;
            this.IsMuted = this.options.StartMuted;

            // 1. Initialize Disk Storage Manager & Timeline
            this.RecordingsManager = new RecordingsManager(this.Call.Id);
            this.AudioAggregator = new AudioAggregator();
            this.Timeline = new MeetingTimeline { CallId = this.Call.Id, ChatThreadId = chatThreadId, StartedAt = this.sessionStartTime };
            this.chatClient = new BotFrameworkChatClient(this.options.AadAppId, this.options.AadAppSecretOrCertThumbprint, this.options.BotFrameworkServiceUrl, this.graphLogger);

            var tdaTokenProvider = new TdaTokenProvider(this.options.Tda, this.options.AadAppId, this.options.AadAppSecretOrCertThumbprint, this.graphLogger);
            this.tdaClient = new TdaClient(this.options.Tda, tdaTokenProvider, this.graphLogger);

            // 2. Wire Call Events
            this.Call.OnUpdated += this.OnCallUpdated;
            this.Call.Participants.OnUpdated += this.OnParticipantsUpdated;

            // 3. Resolve Media Sockets
            var localMediaSession = this.Call.GetLocalMediaSession();
            if (localMediaSession != null)
            {
                // Audio Socket
                this.audioSocket = localMediaSession.AudioSocket;
                if (this.audioSocket != null)
                {
                    this.audioSocket.AudioMediaReceived += this.OnAudioMediaReceived;
                    this.AudioSender = new AudioSender(this.audioSocket, this.graphLogger);
                    this.AudioSender.IsMuted = this.IsMuted;

                    // Play pleasant greeting chime + verbal announcement when audio send is active
                    if (this.options.SpeakGreetingOnJoin && !this.IsMuted)
                    {
                        _ = Task.Run(async () =>
                        {
                            for (int i = 0; i < 30; i++)
                            {
                                if (this.AudioSender != null && this.AudioSender.IsAudioSendActive)
                                {
                                    break;
                                }
                                await Task.Delay(200).ConfigureAwait(false);
                            }

                            if (this.AudioSender != null && this.AudioSender.IsAudioSendActive)
                            {
                                await Task.Delay(500).ConfigureAwait(false);
                                await this.AudioSender.PlayGreetingAsync().ConfigureAwait(false);
                                this.Timeline.AddEvent("bot_greeting_spoken", this.options.GreetingText ?? "Welcome");
                            }
                        });
                    }
                }

                // VBSS (Screen Share Receive) Socket
                this.vbssSocket = localMediaSession.VbssSocket;
                if (this.vbssSocket != null)
                {
                    this.vbssSocket.VideoReceiveStatusChanged += this.OnVbssReceiveStatusChanged;
                    this.vbssSocket.VideoMediaReceived += this.OnVbssMediaReceived;
                    this.vbssSocket.MediaStreamFailure += this.OnMediaStreamFailure;
                    this.Log("[VBSS Socket] Initialised - waiting for a participant to start sharing (subscription happens per presenter MSI).");
                }
                else
                {
                    this.Log("[VBSS Socket] NOT available on this media session - screen share cannot be recorded.");
                }

                // Video Socket (Bot Video Streaming & participant camera receiving)
                if (localMediaSession.VideoSockets != null && localMediaSession.VideoSockets.Any())
                {
                    this.videoSocket = localMediaSession.VideoSockets.First();
                    this.videoSocket.VideoSendStatusChanged += this.OnVideoSendStatusChanged;
                    this.videoSocket.VideoKeyFrameNeeded += this.OnVideoKeyFrameNeeded;
                    this.videoSocket.VideoReceiveStatusChanged += this.OnVideoReceiveStatusChanged;
                    this.videoSocket.VideoMediaReceived += this.OnVideoMediaReceived;
                    this.videoSocket.MediaStreamFailure += this.OnMediaStreamFailure;
                    this.StartBotVideoBroadcast();
                }
            }

            // 4. Start Periodic Audio Disk Flush (every 10s so WAV files are continuously present on disk)
            this.periodicFlushCts = new CancellationTokenSource();
            _ = Task.Run(async () =>
            {
                var token = this.periodicFlushCts.Token;
                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        await Task.Delay(TimeSpan.FromSeconds(10), token).ConfigureAwait(false);
                        this.AudioAggregator.SnapshotToWavFiles(this.RecordingsManager.SessionDirectory);
                    }
                    catch (OperationCanceledException) { }
                    catch (Exception ex)
                    {
                        this.graphLogger.Warn($"Periodic audio snapshot: {ex.Message}");
                    }
                }
            }, this.periodicFlushCts.Token);

            // 5. Hook participants that are already in the roster
            this.HookParticipantUpdates(this.Call.Participants);

            // 6. Start Meeting Chat Monitor
            this.StartMeetingChatMonitor();

            // 7. Start Real-Time Voice Speech Recognizer (replies to spoken voice: "tda bot", "kaise ho", etc.)
            this.voiceRecognizer = new RealtimeSpeechRecognizer(
                this.graphLogger,
                () => this.AudioSender != null && this.AudioSender.IsSpeaking);
            this.voiceRecognizer.OnTriggerDetected += this.OnVoiceTriggerDetected;

            // Diagnostic: log what auth credentials are available
            string diagAppId = BotOptions.Current?.AadAppId ?? "(null)";
            string diagSecret = string.IsNullOrWhiteSpace(BotOptions.Current?.AadAppSecretOrCertThumbprint) ? "(empty)" : $"({BotOptions.Current.AadAppSecretOrCertThumbprint.Length} chars)";
            string diagOverride = string.IsNullOrWhiteSpace(BotOptions.Current?.OverrideBearerToken) ? "(empty)" : $"({BotOptions.Current.OverrideBearerToken.Length} chars)";
            Console.WriteLine($">>> [Config Diagnostic] AppId={diagAppId}, Secret={diagSecret}, OverrideToken={diagOverride}");

            this.graphLogger.Info($"CallHandler initialized for {this.Call.Id}. Output folder: {this.RecordingsManager.SessionDirectory}");
            Console.WriteLine($">>> [CallHandler] Output folder: {this.RecordingsManager.SessionDirectory}");
        }

        private void Log(string message)
        {
            this.graphLogger.Info(message);
            Console.WriteLine(">>> " + message);
            this.RecordingsManager.Log(message);
        }

        // ===================================================================
        // 1. Audio (Both Ways) + Live Voice Triggering
        // ===================================================================
        private void OnAudioMediaReceived(object sender, AudioMediaReceivedEventArgs e)
        {
            try
            {
                var unmixed = e.Buffer.UnmixedAudioBuffers;
                if (unmixed != null && unmixed.Length > 0)
                {
                    foreach (var speakerBuffer in unmixed)
                    {
                        this.AudioAggregator.Append(
                            speakerBuffer.ActiveSpeakerId,
                            speakerBuffer.Data,
                            speakerBuffer.Length,
                            speakerBuffer.OriginalSenderTimestamp);

                        this.voiceRecognizer?.AppendAudio(speakerBuffer.Data, speakerBuffer.Length);
                    }
                }
                else
                {
                    this.AudioAggregator.Append(AudioAggregator.UnknownSpeakerId, e.Buffer.Data, e.Buffer.Length, e.Buffer.Timestamp);
                    this.voiceRecognizer?.AppendAudio(e.Buffer.Data, e.Buffer.Length);
                }
            }
            finally
            {
                e.Buffer.Dispose();
            }
        }

        /// <summary>
        /// Plays an audio file into the Teams meeting so participants can hear the bot.
        /// </summary>
        public async Task PlayAudioFileAsync(string wavFilePath)
        {
            if (this.AudioSender == null) return;
            this.graphLogger.Info($"[Audio Playback] Playing {wavFilePath} into meeting.");
            await this.AudioSender.PlayWavFileAsync(wavFilePath).ConfigureAwait(false);
        }

        // ===================================================================
        // 2. Mute / Unmute
        // ===================================================================
        public async Task MuteAsync()
        {
            this.IsMuted = true;
            if (this.AudioSender != null)
            {
                this.AudioSender.IsMuted = true;
            }

            try
            {
                await this.Call.MuteAsync().ConfigureAwait(false);
                this.graphLogger.Info($"[Mute] Bot muted successfully.");
            }
            catch (Exception ex)
            {
                this.graphLogger.Warn($"[Mute] Server mute returned: {ex.Message}");
            }
        }

        public async Task UnmuteAsync()
        {
            this.IsMuted = false;
            if (this.AudioSender != null)
            {
                this.AudioSender.IsMuted = false;
            }

            try
            {
                await this.Call.UnmuteAsync().ConfigureAwait(false);
                this.graphLogger.Info($"[Unmute] Bot unmuted successfully.");
            }
            catch (Exception ex)
            {
                this.graphLogger.Warn($"[Unmute] Server unmute returned: {ex.Message}");
            }
        }

        // ===================================================================
        // 3. Screen share (VBSS) + camera video: subscription and recording
        // ===================================================================
        private void OnMediaStreamFailure(object sender, MediaStreamFailureEventArgs e)
        {
            this.Log($"[Video] MediaStreamFailure on socket {(sender as IVideoSocket)?.SocketId}: {e}");
        }

        private void OnVbssReceiveStatusChanged(object sender, VideoReceiveStatusChangedEventArgs e)
        {
            this.Log($"[VBSS Socket] VideoReceiveStatusChanged: {e.MediaReceiveStatus}");

            if (e.MediaReceiveStatus == MediaReceiveStatus.Active)
            {
                this.SubscribeParticipantMediaStreams();
            }
            else
            {
                this.StopVbssRecorder("receive status Inactive");
            }
        }

        private void OnVbssMediaReceived(object sender, VideoMediaReceivedEventArgs e)
        {
            try
            {
                if (e.Buffer == null || e.Buffer.Data == IntPtr.Zero)
                {
                    return;
                }

                var recorder = this.vbssRecorder;
                if (recorder != null && this.options.RecordScreenShare)
                {
                    if (e.Buffer.MediaSourceId != 0 && recorder.MediaSourceId != 0 && e.Buffer.MediaSourceId != recorder.MediaSourceId
                        && Interlocked.Exchange(ref this.vbssMsiMismatchLogged, 1) == 0)
                    {
                        this.Log($"[VBSS] Frames arrive with MSI {e.Buffer.MediaSourceId} but recorder was started for MSI {recorder.MediaSourceId} - recording them anyway.");
                    }

                    recorder.OnFrame(e.Buffer); // copies bytes, returns immediately
                    return;
                }

                if (recorder == null && this.options.RecordScreenShare)
                {
                    var who = this.ResolveParticipantByMsi(e.Buffer.MediaSourceId);
                    this.StartVbssRecorder(e.Buffer.MediaSourceId, who.DisplayName, who.ParticipantId);
                    this.vbssRecorder?.OnFrame(e.Buffer);
                    return;
                }

                this.SavePhotoThrottled(e.Buffer, "03_photo_screenshare");
            }
            catch (Exception ex)
            {
                this.graphLogger.Error(ex, "Error handling screen share frame.");
            }
            finally
            {
                e.Buffer?.Dispose();
            }
        }

        private void OnVideoReceiveStatusChanged(object sender, VideoReceiveStatusChangedEventArgs e)
        {
            this.Log($"[Video Socket] VideoReceiveStatusChanged: {e.MediaReceiveStatus}");
            if (e.MediaReceiveStatus == MediaReceiveStatus.Active)
            {
                this.SubscribeParticipantMediaStreams();
            }
            else
            {
                this.StopCameraRecorder("receive status Inactive");
            }
        }

        private void OnVideoMediaReceived(object sender, VideoMediaReceivedEventArgs e)
        {
            try
            {
                if (e.Buffer == null || e.Buffer.Data == IntPtr.Zero)
                {
                    return;
                }

                var recorder = this.cameraRecorder;
                if (recorder != null && this.options.RecordParticipantVideo)
                {
                    recorder.OnFrame(e.Buffer);
                    return;
                }

                if (recorder == null && this.options.RecordParticipantVideo)
                {
                    var who = this.ResolveParticipantByMsi(e.Buffer.MediaSourceId);
                    this.StartCameraRecorder(e.Buffer.MediaSourceId, who.DisplayName, who.ParticipantId);
                    this.cameraRecorder?.OnFrame(e.Buffer);
                    return;
                }

                this.SavePhotoThrottled(e.Buffer, "03_photo_video");
            }
            catch (Exception ex)
            {
                this.graphLogger.Error(ex, "Error handling meeting video frame.");
            }
            finally
            {
                e.Buffer?.Dispose();
            }
        }

        /// <summary>Legacy behaviour (one JPEG every 5 s) - only used when video recording is switched off.</summary>
        private void SavePhotoThrottled(VideoMediaBuffer buffer, string prefix)
        {
            DateTime now = DateTime.Now;
            lock (this.photoLock)
            {
                if (now - this.lastPhotoTime < this.photoInterval)
                {
                    return;
                }

                this.lastPhotoTime = now;
            }

            Bitmap bitmap = null;
            if (buffer.VideoFormat?.VideoColorFormat == VideoColorFormat.NV12)
            {
                bitmap = VideoFrameConverter.ConvertNV12ToBitmap(buffer.Data, buffer.VideoFormat.Width, buffer.VideoFormat.Height, buffer.Stride);
            }
            else if (buffer.VideoFormat?.VideoColorFormat == VideoColorFormat.Rgb24)
            {
                bitmap = VideoFrameConverter.ConvertRGB24ToBitmap(buffer.Data, buffer.VideoFormat.Width, buffer.VideoFormat.Height, buffer.Stride);
            }

            if (bitmap == null)
            {
                return;
            }

            lock (this.photoLock)
            {
                this.latestScreenBitmap?.Dispose();
                this.latestScreenBitmap = (Bitmap)bitmap.Clone();
            }

            var savedPath = this.RecordingsManager.SavePhoto(bitmap, prefix);
            this.Log($"[Photo Saved] {savedPath}");
            bitmap.Dispose();
        }

        /// <summary>
        /// Subscribes VBSS and video sockets to the current presenter / camera.
        /// Filters on stream Direction SendOnly for VBSS; SendOnly/SendReceive for video.
        /// </summary>
        private void SubscribeParticipantMediaStreams()
        {
            try
            {
                if (this.Call?.Participants == null)
                {
                    return;
                }

                string botAppId = this.options.AadAppId;
                uint sharerMsi = 0;
                IParticipant sharer = null;
                uint cameraMsi = 0;
                IParticipant cameraOwner = null;

                foreach (var participant in this.Call.Participants)
                {
                    var resource = participant.Resource;
                    if (resource == null || resource.IsInLobby == true)
                    {
                        continue;
                    }

                    if (resource.Info?.Identity?.Application?.Id == botAppId)
                    {
                        continue;
                    }

                    var streams = resource.MediaStreams;
                    if (streams == null)
                    {
                        continue;
                    }

                    foreach (var stream in streams)
                    {
                        if (!uint.TryParse(stream.SourceId, out uint msi) || msi == 0)
                        {
                            continue;
                        }

                        bool sendCapable = stream.Direction == MediaDirection.SendOnly || stream.Direction == MediaDirection.SendReceive;
                        if (!sendCapable)
                        {
                            continue;
                        }

                        if (stream.MediaType == Modality.VideoBasedScreenSharing && sharer == null)
                        {
                            sharer = participant;
                            sharerMsi = msi;
                        }
                        else if (stream.MediaType == Modality.Video && cameraOwner == null)
                        {
                            cameraOwner = participant;
                            cameraMsi = msi;
                        }
                    }
                }

                // ---- Screen share -------------------------------------------------------------
                if (this.vbssSocket != null)
                {
                    if (sharerMsi != 0 && sharerMsi != this.currentVbssMsi)
                    {
                        string name = GetDisplayName(sharer);
                        try
                        {
                            this.vbssSocket.Subscribe(VideoResolution.HD1080p, sharerMsi);
                            this.currentVbssMsi = sharerMsi;
                            this.Log($"[VBSS Socket] Subscribed to screen share of '{name}' (MSI {sharerMsi}).");
                        }
                        catch (Exception ex)
                        {
                            this.Log($"[VBSS Socket] Subscribe to MSI {sharerMsi} failed: {ex.Message}");
                        }

                        if (this.options.RecordScreenShare)
                        {
                            this.StartVbssRecorder(sharerMsi, name, sharer?.Id);
                        }
                    }
                    else if (sharerMsi == 0 && this.currentVbssMsi != 0)
                    {
                        this.Log($"[VBSS Socket] Presenter (MSI {this.currentVbssMsi}) stopped sharing - unsubscribing.");
                        try
                        {
                            this.vbssSocket.Unsubscribe();
                        }
                        catch (Exception ex)
                        {
                            this.graphLogger.Warn($"[VBSS Socket] Unsubscribe returned: {ex.Message}");
                        }

                        this.currentVbssMsi = 0;
                        this.StopVbssRecorder("presenter stopped sharing");
                    }
                }

                // ---- Camera video ---------------------------------------------------------------
                if (this.videoSocket != null)
                {
                    if (cameraMsi != 0 && cameraMsi != this.currentCameraMsi)
                    {
                        string name = GetDisplayName(cameraOwner);
                        try
                        {
                            this.videoSocket.Subscribe(VideoResolution.HD720p, cameraMsi);
                            this.currentCameraMsi = cameraMsi;
                            this.Log($"[Video Socket] Subscribed to camera of '{name}' (MSI {cameraMsi}).");
                        }
                        catch (Exception ex)
                        {
                            this.Log($"[Video Socket] Subscribe to MSI {cameraMsi} failed: {ex.Message}");
                        }

                        if (this.options.RecordParticipantVideo)
                        {
                            this.StartCameraRecorder(cameraMsi, name, cameraOwner?.Id);
                        }
                    }
                    else if (cameraMsi == 0 && this.currentCameraMsi != 0)
                    {
                        this.Log($"[Video Socket] Camera (MSI {this.currentCameraMsi}) switched off - unsubscribing.");
                        try
                        {
                            this.videoSocket.Unsubscribe();
                        }
                        catch (Exception ex)
                        {
                            this.graphLogger.Warn($"[Video Socket] Unsubscribe returned: {ex.Message}");
                        }

                        this.currentCameraMsi = 0;
                        this.StopCameraRecorder("camera switched off");
                    }
                }
            }
            catch (Exception ex)
            {
                this.graphLogger.Warn($"Error subscribing participant media streams: {ex.Message}");
            }
        }

        private VideoRecorderSettings ScreenShareRecorderSettings() => new VideoRecorderSettings
        {
            Fps = Math.Max(1, this.options.ScreenShareRecordingFps),
            JpegQuality = this.options.VideoJpegQuality,
            SnapshotIntervalSeconds = Math.Max(1, this.options.SnapshotIntervalSeconds),
            MaxSegmentBytes = Math.Max(50_000_000, this.options.MaxVideoSegmentBytes),
            SnapshotPrefix = "03_snapshot_screenshare",
        };

        private VideoRecorderSettings CameraRecorderSettings() => new VideoRecorderSettings
        {
            Fps = Math.Max(1, this.options.ParticipantVideoRecordingFps),
            JpegQuality = this.options.VideoJpegQuality,
            SnapshotIntervalSeconds = Math.Max(5, this.options.SnapshotIntervalSeconds * 3),
            MaxSegmentBytes = Math.Max(50_000_000, this.options.MaxVideoSegmentBytes),
            SnapshotPrefix = "03_snapshot_camera",
        };

        private void StartVbssRecorder(uint msi, string presenterName, string participantId)
        {
            lock (this.videoLock)
            {
                if (this.vbssRecorder != null && this.vbssRecorder.MediaSourceId == msi)
                {
                    return;
                }

                this.StopVbssRecorderCore("new presenter");

                string stem = $"03_video_screenshare_{VideoRecorder.SanitizeFileName(presenterName)}_{DateTime.Now:HHmmss}";
                var recorder = new VideoRecorder(this.RecordingsManager.SessionDirectory, stem, msi, presenterName, this.ScreenShareRecorderSettings(), this.RecordingsManager.Log);
                recorder.SnapshotSaved += (r, snap) => this.graphLogger.Info($"[Screen Snapshot Saved] {snap.Path}");
                recorder.Start();

                this.vbssRecorder = recorder;
                this.vbssSession = this.Timeline.StartScreenShare(msi, presenterName, participantId);
                this.activityLine = $"Recording {presenterName}'s screen share";
                Interlocked.Exchange(ref this.vbssMsiMismatchLogged, 0);
            }

            this.Log($"[Screen Recording] Started for '{presenterName}' (MSI {msi}).");

            if (this.options.AnnounceScreenShare)
            {
                _ = Task.Run(async () =>
                {
                    try
                    {
                        _ = this.PostTextMessageToChatAsync($"🎥 <b>Recording started:</b> {System.Net.WebUtility.HtmlEncode(presenterName)}'s screen share is now being recorded.");
                        if (this.AudioSender != null && !this.IsMuted)
                        {
                            await this.AudioSender.SpeakAsync($"I have started recording {presenterName}'s screen share.").ConfigureAwait(false);
                        }
                    }
                    catch (Exception ex)
                    {
                        this.graphLogger.Warn($"[Screen Recording] Announcement failed: {ex.Message}");
                    }
                });
            }
        }

        private void StopVbssRecorder(string reason)
        {
            lock (this.videoLock)
            {
                this.StopVbssRecorderCore(reason);
            }
        }

        private void StopVbssRecorderCore(string reason)
        {
            var recorder = this.vbssRecorder;
            if (recorder == null)
            {
                return;
            }

            this.vbssRecorder = null;
            this.activityLine = null;
            try
            {
                recorder.Stop();
                this.Timeline.EndMediaSession(this.vbssSession, recorder.SegmentPaths, recorder.Snapshots.Select(s => s.Path), recorder.FramesWritten);
                this.Log($"[Screen Recording] Stopped for '{recorder.SourceLabel}' ({reason}): {recorder.FramesWritten} frames written from {recorder.FramesReceived} received -> {string.Join(", ", recorder.SegmentPaths.Select(Path.GetFileName))}");
            }
            catch (Exception ex)
            {
                this.graphLogger.Error(ex, "[Screen Recording] Error while stopping recorder.");
            }
            finally
            {
                recorder.Dispose();
                this.vbssSession = null;
            }
        }

        private void StartCameraRecorder(uint msi, string participantName, string participantId)
        {
            lock (this.videoLock)
            {
                if (this.cameraRecorder != null && this.cameraRecorder.MediaSourceId == msi)
                {
                    return;
                }

                this.StopCameraRecorderCore("new camera source");

                string stem = $"03_video_camera_{VideoRecorder.SanitizeFileName(participantName)}_{DateTime.Now:HHmmss}";
                var recorder = new VideoRecorder(this.RecordingsManager.SessionDirectory, stem, msi, participantName, this.CameraRecorderSettings(), this.RecordingsManager.Log);
                recorder.Start();
                this.cameraRecorder = recorder;
                this.cameraSession = this.Timeline.StartCameraSession(msi, participantName, participantId);
            }

            this.Log($"[Camera Recording] Started for '{participantName}' (MSI {msi}).");
        }

        private void StopCameraRecorder(string reason)
        {
            lock (this.videoLock)
            {
                this.StopCameraRecorderCore(reason);
            }
        }

        private void StopCameraRecorderCore(string reason)
        {
            var recorder = this.cameraRecorder;
            if (recorder == null)
            {
                return;
            }

            this.cameraRecorder = null;
            try
            {
                recorder.Stop();
                this.Timeline.EndMediaSession(this.cameraSession, recorder.SegmentPaths, recorder.Snapshots.Select(s => s.Path), recorder.FramesWritten);
                this.Log($"[Camera Recording] Stopped for '{recorder.SourceLabel}' ({reason}): {recorder.FramesWritten} frames -> {string.Join(", ", recorder.SegmentPaths.Select(Path.GetFileName))}");
            }
            catch (Exception ex)
            {
                this.graphLogger.Error(ex, "[Camera Recording] Error while stopping recorder.");
            }
            finally
            {
                recorder.Dispose();
                this.cameraSession = null;
            }
        }

        /// <summary>
        /// Instantly saves a photo of the current active screen share (or meeting video) on demand.
        /// </summary>
        public string CapturePhotoNow(string tag = "manual")
        {
            var recorder = this.vbssRecorder ?? this.cameraRecorder;
            if (recorder != null)
            {
                string path = Path.Combine(this.RecordingsManager.SessionDirectory, $"03_photo_{tag}_{DateTime.Now:yyyyMMdd_HHmmss_fff}.jpg");
                var saved = recorder.SaveLatestFrame(path);
                if (saved != null)
                {
                    this.RecordingsManager.Log($"[Photo Saved] {saved}");
                    return saved;
                }
            }

            lock (this.photoLock)
            {
                if (this.latestScreenBitmap != null)
                {
                    return this.RecordingsManager.SavePhoto(this.latestScreenBitmap, $"03_photo_{tag}");
                }
            }

            return null;
        }

        /// <summary>Summary of the active/finished video recordings for the management API.</summary>
        public object GetVideoStatus()
        {
            var vbss = this.vbssRecorder;
            var cam = this.cameraRecorder;
            return new
            {
                ScreenShare = vbss == null ? null : new
                {
                    Presenter = vbss.SourceLabel,
                    Msi = vbss.MediaSourceId,
                    StartedAt = vbss.StartedAt,
                    Resolution = $"{vbss.Width}x{vbss.Height}",
                    FramesReceived = vbss.FramesReceived,
                    FramesWritten = vbss.FramesWritten,
                    Files = vbss.SegmentPaths,
                    Snapshots = vbss.Snapshots.Count,
                },
                Camera = cam == null ? null : new
                {
                    Participant = cam.SourceLabel,
                    Msi = cam.MediaSourceId,
                    StartedAt = cam.StartedAt,
                    Resolution = $"{cam.Width}x{cam.Height}",
                    FramesReceived = cam.FramesReceived,
                    FramesWritten = cam.FramesWritten,
                    Files = cam.SegmentPaths,
                },
                CompletedScreenShareSessions = this.Timeline.ScreenShareSessions.Where(s => s.EndedAt != null).Select(s => new { s.PresenterName, s.StartedAt, s.EndedAt, s.VideoFiles, SnapshotCount = s.SnapshotFiles.Count }),
                BotVideoSendActive = this.isVideoSendActive,
                BotVideoFormat = this.preferredSendFormat?.ToString(),
            };
        }

        // ===================================================================
        // 5. Video Sharing by Bot (Status Card Broadcast)
        // ===================================================================
        private void OnVideoSendStatusChanged(object sender, VideoSendStatusChangedEventArgs e)
        {
            var preferred = e.PreferredVideoSourceFormat;
            if (preferred != null && preferred.VideoColorFormat == VideoColorFormat.NV12 && preferred.Width > 0 && preferred.Height > 0)
            {
                this.preferredSendFormat = preferred;
            }

            this.isVideoSendActive = (e.MediaSendStatus == MediaSendStatus.Active);
            this.Log($"[VideoSocket] VideoSendStatusChanged: {e.MediaSendStatus}, preferred format: {preferred}");
        }

        private void OnVideoKeyFrameNeeded(object sender, VideoKeyFrameNeededEventArgs e)
        {
            this.graphLogger.Info("[VideoSocket] KeyFrameNeeded");
        }

        private void StartBotVideoBroadcast()
        {
            if (this.videoSocket == null) return;

            this.videoBroadcastCts = new CancellationTokenSource();
            var token = this.videoBroadcastCts.Token;

            _ = Task.Run(async () =>
            {
                int width = 1280;
                int height = 720;
                int tick = 0;
                bool snapshotSaved = false;
                var pace = Stopwatch.StartNew();

                // CRITICAL STABILITY: Pre-allocated ring buffer pool with SafeVideoMediaBuffer
                // avoids 0xc0000005 crash from premature Marshal.FreeHGlobal while native encoder is reading.
                const int poolSize = 6;
                int maxBufSize = width * height * 3 / 2;
                IntPtr[] videoBufferPool = new IntPtr[poolSize];
                for (int i = 0; i < poolSize; i++)
                {
                    videoBufferPool[i] = Marshal.AllocHGlobal(maxBufSize);
                }

                try
                {
                    while (!token.IsCancellationRequested)
                    {
                        tick++;
                        var format = this.preferredSendFormat ?? VideoFormat.NV12_1280x720_15Fps;
                        int currentW = format.Width > 0 ? format.Width : width;
                        int currentH = format.Height > 0 ? format.Height : height;
                        int bufSize = currentW * currentH * 3 / 2;
                        int frameDelayMs = (int)Math.Max(33, 1000 / Math.Max(1f, format.FrameRate));

                        try
                        {
                            if (!snapshotSaved)
                            {
                                snapshotSaved = true;
                                using (var card = this.RenderCurrentBotFrame(tick))
                                {
                                    this.RecordingsManager.SavePhoto(card, "05_bot_broadcast_preview");
                                }
                            }

                            if (this.isVideoSendActive && bufSize <= maxBufSize)
                            {
                                using (var card = this.RenderCurrentBotFrame(tick))
                                {
                                    byte[] nv12Bytes = VideoFrameConverter.ConvertBitmapToNV12(card, currentW, currentH);
                                    int slot = tick % poolSize;
                                    Marshal.Copy(nv12Bytes, 0, videoBufferPool[slot], nv12Bytes.Length);

                                    var videoBuffer = new SafeVideoMediaBuffer(
                                        videoBufferPool[slot],
                                        nv12Bytes.Length,
                                        format,
                                        MediaPlatform.GetCurrentTimestamp());

                                    this.videoSocket.Send(videoBuffer);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (Interlocked.Increment(ref this.broadcastErrorsLogged) <= 5)
                            {
                                this.graphLogger.Warn($"[Bot Video] Frame send failed (loop continues): {ex.Message}");
                            }
                        }

                        long elapsed = pace.ElapsedMilliseconds;
                        int wait = (int)Math.Max(1, frameDelayMs - elapsed);
                        pace.Restart();
                        try
                        {
                            await Task.Delay(wait, token).ConfigureAwait(false);
                        }
                        catch (OperationCanceledException)
                        {
                            break;
                        }
                    }
                }
                finally
                {
                    try { await Task.Delay(400).ConfigureAwait(false); } catch { }
                    for (int i = 0; i < poolSize; i++)
                    {
                        if (videoBufferPool[i] != IntPtr.Zero)
                        {
                            Marshal.FreeHGlobal(videoBufferPool[i]);
                            videoBufferPool[i] = IntPtr.Zero;
                        }
                    }
                }

                this.graphLogger.Info("[Bot Video Streaming] Broadcast loop stopped.");
            }, token);

            this.graphLogger.Info("[Bot Video Streaming] Broadcast loop initialized at 15 FPS.");
        }

        /// <summary>
        /// Picks between the normal status card and an active visualization (chart / TDA answer /
        /// image shown on the bot's own video tile - see ShowVisualization). Caller owns and
        /// must Dispose the returned Bitmap.
        /// </summary>
        private Bitmap RenderCurrentBotFrame(int tick)
        {
            Bitmap visualization = null;
            string title = null;
            lock (this.visualizationLock)
            {
                if (this.currentVisualization != null && DateTime.Now < this.visualizationExpiresAt)
                {
                    visualization = this.currentVisualization;
                    title = this.currentVisualizationTitle;
                }
                else if (this.currentVisualization != null)
                {
                    // Expired - clear it so subsequent frames fall back to the status card.
                    this.currentVisualization.Dispose();
                    this.currentVisualization = null;
                    this.currentVisualizationTitle = null;
                }
            }

            if (visualization != null)
            {
                return VideoFrameConverter.CreateVisualizationFrame(visualization, title);
            }

            return VideoFrameConverter.CreateBotStatusCard(
                "Teams AI Assistant",
                this.IsMuted ? "Audio output muted - still recording" : "Recording audio and video",
                this.Call.Id,
                this.IsMuted,
                tick,
                this.activityLine);
        }

        /// <summary>
        /// Shows <paramref name="contentImage"/> on the bot's outgoing video tile for
        /// <paramref name="durationSeconds"/> seconds (0 or negative = show indefinitely until
        /// ClearVisualization or a new call to this method). Takes ownership of contentImage (clones
        /// it internally, caller may dispose its own copy). This is the "share screen to show
        /// visualisation" capability.
        /// </summary>
        public void ShowVisualization(Bitmap contentImage, string title, int durationSeconds)
        {
            if (contentImage == null)
            {
                return;
            }

            var clone = (Bitmap)contentImage.Clone();
            lock (this.visualizationLock)
            {
                this.currentVisualization?.Dispose();
                this.currentVisualization = clone;
                this.currentVisualizationTitle = title;
                this.visualizationExpiresAt = durationSeconds > 0
                    ? DateTime.Now.AddSeconds(durationSeconds)
                    : DateTime.MaxValue;
            }

            this.Log($"[Visualization] Showing '{title}' on bot video tile" + (durationSeconds > 0 ? $" for {durationSeconds}s." : " indefinitely."));
        }

        /// <summary>Reverts the bot's video tile to the normal status card immediately.</summary>
        public void ClearVisualization()
        {
            lock (this.visualizationLock)
            {
                this.currentVisualization?.Dispose();
                this.currentVisualization = null;
                this.currentVisualizationTitle = null;
                this.visualizationExpiresAt = DateTime.MinValue;
            }
        }

        // ===================================================================
        // TDA (Tata Steel Digital Assistant) integration - inert unless Bot:Tda:Enabled is true.
        // See Config/BotOptions.cs (TdaOptions) and Tda/TdaClient.cs.
        // ===================================================================

        /// <summary>
        /// Asks TDA <paramref name="query"/> and speaks the answer into the meeting. No-ops (logs and
        /// returns false) if TDA is not configured/enabled, if AudioSender is unavailable, or if TDA
        /// returns nothing - never throws into the caller.
        /// </summary>
        public async Task<bool> SpeakTdaAnswerAsync(string query)
        {
            if (!this.tdaClient.IsConfigured)
            {
                this.graphLogger.Warn("[TDA] SpeakTdaAnswerAsync called but TDA is not enabled/configured - no-op.");
                return false;
            }

            try
            {
                string answer = await this.tdaClient.AskAsync(query).ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(answer))
                {
                    return false;
                }

                if (this.AudioSender != null && !this.IsMuted)
                {
                    await this.AudioSender.SpeakAsync(answer).ConfigureAwait(false);
                }

                this.Timeline.AddEvent("tda_answer_spoken", answer);
                return true;
            }
            catch (Exception ex)
            {
                this.graphLogger.Warn($"[TDA] SpeakTdaAnswerAsync failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>Sends a message to TDA using a token scoped to the "TSL AI" resource. False/no-op if not configured.</summary>
        public async Task<bool> SendTdaMessageAsync(string message)
        {
            if (!this.tdaClient.IsConfigured)
            {
                this.graphLogger.Warn("[TDA] SendTdaMessageAsync called but TDA is not enabled/configured - no-op.");
                return false;
            }

            try
            {
                return await this.tdaClient.SendMessageAsync(message).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.graphLogger.Warn($"[TDA] SendTdaMessageAsync failed: {ex.Message}");
                return false;
            }
        }

        // ===================================================================
        // 6. Participants: roster tracking, greetings, auto leave
        // ===================================================================
        private void OnParticipantsUpdated(IParticipantCollection sender, CollectionEventArgs<IParticipant> args)
        {
            string botAppId = this.options.AadAppId;

            foreach (var added in args.AddedResources)
            {
                bool isBot = added.Resource?.Info?.Identity?.Application?.Id == botAppId;
                string name = GetDisplayName(added);
                this.Timeline.ParticipantJoined(added.Id, name, added.Resource?.Info?.Identity?.User?.Id, isBot);

                if (!isBot && added.Resource?.IsInLobby != true && this.options.GreetParticipantsByName && this.callEstablished
                    && (DateTime.Now - this.sessionStartTime).TotalSeconds > 20)
                {
                    var greetName = name;
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await Task.Delay(2500).ConfigureAwait(false);
                            if (this.AudioSender != null && !this.IsMuted)
                            {
                                await this.AudioSender.SpeakAsync($"Hi {FirstName(greetName)}, welcome to the meeting.").ConfigureAwait(false);
                            }
                        }
                        catch { }
                    });
                }
            }

            foreach (var removed in args.RemovedResources)
            {
                this.Timeline.ParticipantLeft(removed.Id);
                lock (this.participantsWithUpdateHook)
                {
                    if (this.participantsWithUpdateHook.Remove(removed.Id))
                    {
                        try { removed.OnUpdated -= this.OnParticipantUpdated; } catch { }
                    }
                }
            }

            this.HookParticipantUpdates(args.AddedResources);

            int humanCount = this.Call.Participants.Count(p =>
                p.Resource?.IsInLobby == false &&
                (p.Resource.Info?.Identity?.Application == null || p.Resource.Info.Identity.Application.Id != botAppId));

            this.graphLogger.Info($"[Participants Updated] Active non-bot participant count: {humanCount}");
            Console.WriteLine($">>> [Participants Updated] Active non-bot participant count: {humanCount}");

            this.SubscribeParticipantMediaStreams();

            bool inGracePeriod = (DateTime.Now - this.sessionStartTime).TotalSeconds < InitialGracePeriodSeconds;

            if (humanCount == 0 && !inGracePeriod)
            {
                if (this.autoLeaveCts == null || this.autoLeaveCts.IsCancellationRequested)
                {
                    this.autoLeaveCts = new CancellationTokenSource();
                    var token = this.autoLeaveCts.Token;

                    this.graphLogger.Warn($"[Auto Leave] All participants left. Leaving in {AutoLeaveDebounceSeconds}s...");
                    Console.WriteLine($">>> [Auto Leave] All participants left. Bot leaving in {AutoLeaveDebounceSeconds} seconds...");

                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await Task.Delay(TimeSpan.FromSeconds(AutoLeaveDebounceSeconds), token).ConfigureAwait(false);
                            if (!token.IsCancellationRequested)
                            {
                                Console.WriteLine(">>> [Auto Leave] Hanging up call now.");
                                this.RecordingsManager.SaveAutoLeaveRecord("All human participants left the meeting", 0);
                                await this.Call.DeleteAsync().ConfigureAwait(false);
                            }
                        }
                        catch (OperationCanceledException)
                        {
                            Console.WriteLine(">>> [Auto Leave] Countdown canceled - participant returned.");
                        }
                        catch (Exception ex)
                        {
                            this.graphLogger.Error(ex, "Failed executing Auto Leave call deletion.");
                        }
                    }, token);
                }
            }
            else if (humanCount > 0)
            {
                if (this.autoLeaveCts != null && !this.autoLeaveCts.IsCancellationRequested)
                {
                    this.autoLeaveCts.Cancel();
                    this.autoLeaveCts = null;
                    this.graphLogger.Info("[Auto Leave] Canceled because a human participant is present.");
                    Console.WriteLine(">>> [Auto Leave] Countdown canceled - participant is present.");
                }
            }
        }

        private void HookParticipantUpdates(IEnumerable<IParticipant> participants)
        {
            if (participants == null) return;

            foreach (var participant in participants)
            {
                lock (this.participantsWithUpdateHook)
                {
                    if (!this.participantsWithUpdateHook.Add(participant.Id))
                    {
                        continue;
                    }
                }

                participant.OnUpdated += this.OnParticipantUpdated;
            }
        }

        private void OnParticipantUpdated(IParticipant sender, ResourceEventArgs<Participant> args)
        {
            try
            {
                var streams = args.NewResource?.MediaStreams;
                if (streams != null)
                {
                    var summary = string.Join(", ", streams.Select(s => $"{s.MediaType}:{s.Direction}:{s.SourceId}"));
                    this.graphLogger.Info($"[Participant Updated] {GetDisplayName(sender)} streams -> {summary}");
                }
            }
            catch { }

            this.SubscribeParticipantMediaStreams();
        }

        private static string GetDisplayName(IParticipant participant)
        {
            var identity = participant?.Resource?.Info?.Identity;
            var name = identity?.User?.DisplayName;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = identity?.Application?.DisplayName;
            }

            if (string.IsNullOrWhiteSpace(name) && identity?.AdditionalData != null)
            {
                foreach (var kvp in identity.AdditionalData)
                {
                    var token = kvp.Value as Newtonsoft.Json.Linq.JObject;
                    var dn = token?["displayName"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(dn))
                    {
                        name = dn;
                        break;
                    }
                }
            }

            return string.IsNullOrWhiteSpace(name) ? $"Participant {participant?.Id?.Substring(0, Math.Min(8, participant.Id.Length))}" : name;
        }

        private static string FirstName(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName)) return "there";
            var parts = displayName.Trim().Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length > 0 ? parts[0] : displayName;
        }

        private (string ParticipantId, string DisplayName) ResolveParticipantByMsi(uint msi)
        {
            try
            {
                var participant = this.Call.Participants.FirstOrDefault(p =>
                    p.Resource?.MediaStreams != null &&
                    p.Resource.MediaStreams.Any(m => m.SourceId == msi.ToString()));

                if (participant != null)
                {
                    return (participant.Id, GetDisplayName(participant));
                }
            }
            catch { }

            return (null, msi == 0 ? "Presenter" : $"Presenter_MSI{msi}");
        }

        // ===================================================================
        // 7. Call state, removal detection & chat message posting
        // ===================================================================
        private void OnCallUpdated(ICall sender, ResourceEventArgs<Call> args)
        {
            Console.WriteLine($">>> CALL STATE CHANGED: {args.NewResource.State} (call {this.Call.Id})");
            var resultInfo = args.NewResource.ResultInfo;
            if (resultInfo != null)
            {
                Console.WriteLine($">>> CALL RESULT INFO: code={resultInfo.Code}, subcode={resultInfo.Subcode}, message={resultInfo.Message}");
            }

            if (args.NewResource.State == CallState.Established)
            {
                this.callEstablished = true;
                if (Interlocked.Exchange(ref this.joinWelcomeSent, 1) == 0)
                {
                    this.Timeline.AddEvent("call_established", this.Call.Id);
                    _ = Task.Run(() => this.PostTextMessageToChatAsync(
                        "🤖 <b>Teams AI Assistant</b> has joined the meeting.<br/>• Audio recording &amp; speech responses: <b>Active</b><br/>• Screen share video recording: <b>Ready</b><br/>• Bot video status tile: <b>Active</b><br/>• Automatic Minutes of Meeting (Word doc): <b>Enabled</b>"));
                }

                this.SubscribeParticipantMediaStreams();
            }

            if (args.NewResource.State != CallState.Terminated)
            {
                return;
            }

            if (Interlocked.Exchange(ref this.endHandled, 1) == 1)
            {
                return;
            }

            bool wasRemovedByUser = false;
            if (resultInfo != null)
            {
                string msg = resultInfo.Message?.ToLowerInvariant() ?? "";
                if (resultInfo.Subcode == 7002 || resultInfo.Subcode == 702 || resultInfo.Code == 403 ||
                    msg.Contains("removed") || msg.Contains("kicked") || msg.Contains("ejected") || msg.Contains("organizer"))
                {
                    wasRemovedByUser = true;
                }
            }

            if (wasRemovedByUser)
            {
                Console.WriteLine(">>> BOT REMOVAL DETECTED: Announcing and posting notification message...");
                _ = Task.Run(async () =>
                {
                    try
                    {
                        if (this.AudioSender != null)
                        {
                            await this.AudioSender.SpeakAsync("Teams AI Assistant has been removed from this meeting.").ConfigureAwait(false);
                        }
                    }
                    catch { }
                });

                _ = Task.Run(() => this.PostTextMessageToChatAsync(
                    "⚠️ <b>Teams Calling Bot Notification:</b> The bot was removed from this meeting by a participant or organizer.",
                    isRemovalNotification: true));
            }
            else
            {
                Console.WriteLine(">>> CALL TERMINATION / BOT LEAVING: Posting group chat departure notification...");
                _ = Task.Run(() => this.PostTextMessageToChatAsync(
                    "👋 <b>Teams Calling Bot Notification:</b> The bot has left the meeting. All session audio, video recordings, transcripts, and Minutes of Meeting have been archived.",
                    isRemovalNotification: true));
            }

            _ = this.HandleCallEndedAsync().ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    this.graphLogger.Error(t.Exception, $"Failed during wind-down for call {this.Call.Id}.");
                }
            });
        }

        public string GetEffectiveChatThreadId()
        {
            if (!string.IsNullOrWhiteSpace(this.chatThreadId))
            {
                return this.chatThreadId;
            }

            try
            {
                var threadId = this.Call.Resource?.ChatInfo?.ThreadId;
                if (!string.IsNullOrWhiteSpace(threadId))
                {
                    return threadId;
                }
            }
            catch { }

            return null;
        }

        public string GetEffectiveAccessToken()
        {
            if (!string.IsNullOrWhiteSpace(this.botAccessToken))
            {
                return this.botAccessToken;
            }

            return BotOptions.Current?.OverrideBearerToken;
        }

        private static string cachedBfToken = null;
        private static DateTime bfTokenExpiry = DateTime.MinValue;
        private static readonly object bfTokenLock = new object();

        public static async Task<string> GetBotFrameworkTokenAsync(IGraphLogger logger = null)
        {
            lock (bfTokenLock)
            {
                if (!string.IsNullOrWhiteSpace(cachedBfToken) && DateTime.UtcNow < bfTokenExpiry)
                {
                    return cachedBfToken;
                }
            }

            BotOptions.Reload();
            string appId = BotOptions.Current?.AadAppId;
            string appSecret = BotOptions.Current?.AadAppSecretOrCertThumbprint;
            string tenantId = BotOptions.Current?.AadTenantId;

            if (string.IsNullOrWhiteSpace(appId) || string.IsNullOrWhiteSpace(appSecret))
            {
                string missing = $"AppId={(string.IsNullOrWhiteSpace(appId) ? "MISSING" : "OK")}, Secret={(string.IsNullOrWhiteSpace(appSecret) ? "MISSING" : "OK")}";
                Console.WriteLine($">>> [BF Token] CANNOT acquire token: {missing}");
                logger?.Warn($"[BF Token] CANNOT acquire token: {missing}");
                return null;
            }

            var tokenEndpoints = !string.IsNullOrWhiteSpace(tenantId) && !tenantId.Equals("common", StringComparison.OrdinalIgnoreCase)
                ? new[] { $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token" }
                : new[] { "https://login.microsoftonline.com/common/oauth2/v2.0/token" };

            foreach (var tokenEndpoint in tokenEndpoints)
            {
                try
                {
                    using (var client = new HttpClient())
                    {
                        var pairs = new List<KeyValuePair<string, string>>
                        {
                            new KeyValuePair<string, string>("grant_type", "client_credentials"),
                            new KeyValuePair<string, string>("client_id", appId),
                            new KeyValuePair<string, string>("client_secret", appSecret),
                            new KeyValuePair<string, string>("scope", "https://api.botframework.com/.default")
                        };

                        var content = new FormUrlEncodedContent(pairs);
                        var resp = await client.PostAsync(tokenEndpoint, content).ConfigureAwait(false);
                        var json = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);

                        if (resp.IsSuccessStatusCode)
                        {
                            var jobj = JObject.Parse(json);
                            string token = jobj["access_token"]?.ToString();
                            int expiresIn = jobj["expires_in"]?.Value<int>() ?? 3600;

                            lock (bfTokenLock)
                            {
                                cachedBfToken = token;
                                bfTokenExpiry = DateTime.UtcNow.AddSeconds(expiresIn - 60);
                            }

                            return token;
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger?.Error(ex, $"[BF Token] EXCEPTION on {tokenEndpoint}");
                }
            }

            return null;
        }

        public async Task<bool> PostActivityViaBotFrameworkAsync(string threadId, string text)
        {
            string bfToken = await GetBotFrameworkTokenAsync(this.graphLogger).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(bfToken))
            {
                return false;
            }

            string botAppId = BotOptions.Current?.AadAppId ?? string.Empty;
            string tenantId = BotOptions.Current?.AadTenantId ?? string.Empty;

            var serviceUrls = new[]
            {
                "https://smba.trafficmanager.net/apis/",
                "https://smba.trafficmanager.net/apac/",
                "https://smba.trafficmanager.net/amer/",
                "https://smba.trafficmanager.net/emea/"
            };

            foreach (var serviceUrl in serviceUrls)
            {
                string url = $"{serviceUrl}v3/conversations/{threadId}/activities";
                var activity = new
                {
                    type = "message",
                    text = text,
                    textFormat = "markdown",
                    channelId = "msteams",
                    serviceUrl = serviceUrl,
                    from = new { id = botAppId, name = "TDA Bot" },
                    conversation = new { id = threadId, isGroup = true, tenantId = tenantId },
                    channelData = new { tenant = new { id = tenantId } }
                };

                string jsonPayload = JsonConvert.SerializeObject(activity);

                try
                {
                    using (var client = new HttpClient())
                    {
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bfToken);
                        var reqContent = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                        var resp = await client.PostAsync(url, reqContent).ConfigureAwait(false);

                        if (resp.IsSuccessStatusCode)
                        {
                            this.graphLogger.Info($"[BotFramework Connector] Successfully posted message to {serviceUrl}");
                            Console.WriteLine($">>> [BotFramework Connector] Message sent successfully via {serviceUrl}");
                            return true;
                        }
                    }
                }
                catch { }
            }

            return false;
        }

        /// <summary>
        /// Posts a formatted text/HTML message into the Teams meeting group chat.
        /// Primary: BotFrameworkChatClient connector -> Fallback: direct BF API -> Fallback: Graph chat API.
        /// </summary>
        public async Task<bool> PostTextMessageToChatAsync(string htmlContent, bool isRemovalNotification = false)
        {
            string threadId = this.GetEffectiveChatThreadId();
            if (string.IsNullOrWhiteSpace(threadId))
            {
                this.graphLogger.Warn("[Chat Notification] Missing chatThreadId - cannot post message.");
                return false;
            }

            if (this.options.UseBotFrameworkForChat && this.chatClient.IsConfigured)
            {
                bool viaConnector = await this.chatClient.SendMessageAsync(threadId, htmlContent).ConfigureAwait(false);
                if (viaConnector)
                {
                    this.Timeline.AddEvent("chat_posted", "connector");
                    if (isRemovalNotification)
                    {
                        this.RecordingsManager.SaveRemovalMessageRecord(threadId, htmlContent, true);
                    }
                    return true;
                }

                this.graphLogger.Warn("[Chat Notification] Connector post failed - falling back to direct Bot Framework / Graph chat API.");
            }

            string markdownText = htmlContent
                .Replace("<b>", "**").Replace("</b>", "**")
                .Replace("<i>", "*").Replace("</i>", "*")
                .Replace("<br/>", "\n").Replace("<br>", "\n");

            bool bfSuccess = await this.PostActivityViaBotFrameworkAsync(threadId, markdownText).ConfigureAwait(false);
            if (bfSuccess)
            {
                this.Timeline.AddEvent("chat_posted", "botframework");
                if (isRemovalNotification)
                {
                    this.RecordingsManager.SaveRemovalMessageRecord(threadId, htmlContent, true);
                }
                return true;
            }

            string token = this.GetEffectiveAccessToken();
            if (string.IsNullOrWhiteSpace(token))
            {
                this.graphLogger.Warn("[Chat Notification] No access token available to post chat message.");
                if (isRemovalNotification)
                {
                    this.RecordingsManager.SaveRemovalMessageRecord(threadId, htmlContent, false);
                }
                return false;
            }

            try
            {
                string endpoint = $"https://graph.microsoft.com/v1.0/chats/{threadId}/messages";

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    var payload = new
                    {
                        body = new
                        {
                            contentType = "html",
                            content = htmlContent
                        }
                    };

                    string jsonContent = JsonConvert.SerializeObject(payload);
                    var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(endpoint, content).ConfigureAwait(false);
                    bool success = response.IsSuccessStatusCode;

                    if (isRemovalNotification)
                    {
                        this.RecordingsManager.SaveRemovalMessageRecord(threadId, htmlContent, success);
                    }

                    if (success)
                    {
                        this.graphLogger.Info($"[Chat Notification] Successfully posted message to meeting chat {threadId}!");
                        Console.WriteLine($">>> [Chat Notification] Successfully posted message to meeting chat!");
                        this.Timeline.AddEvent("chat_posted", "graph");
                        return true;
                    }
                    else
                    {
                        var err = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        this.graphLogger.Warn($"[Chat Notification] Graph chat post returned ({response.StatusCode}): {err}");
                        Console.WriteLine($">>> [Chat Notification] Graph chat post returned ({response.StatusCode}): {err}");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                this.graphLogger.Error(ex, "Error posting message to meeting chat.");
                if (isRemovalNotification)
                {
                    this.RecordingsManager.SaveRemovalMessageRecord(threadId, $"Exception: {ex.Message}", false);
                }
                return false;
            }
        }

        private void StartMeetingChatMonitor()
        {
            this.chatMonitorCts = new CancellationTokenSource();
            string threadId = this.GetEffectiveChatThreadId();
            Console.WriteLine($">>> [Chat Monitor] Initialized. Meeting chat thread: {threadId ?? "(listening via /api/messages)"}.");
        }

        public void HandleIncomingChatActivity(string fromUserName, string fromUserId, string messageText, string incomingServiceUrl)
        {
            if (string.IsNullOrWhiteSpace(messageText)) return;

            string lower = messageText.ToLowerInvariant();

            bool isBotMention = lower.Contains("tda") || lower.Contains("bot") ||
                                 lower.Contains("kaise ho") || lower.Contains("kaisa") ||
                                 lower.Contains("hi") || lower.Contains("hello") ||
                                 lower.Contains("hey") || lower.Contains("help") ||
                                 lower.Contains("status") || lower.Contains("recording");

            if (!isBotMention) return;

            this.graphLogger.Info($"[Chat Monitor] Incoming message from {fromUserName}: \"{messageText}\"");
            Console.WriteLine($">>> [Chat Monitor] Incoming message from {fromUserName}: \"{messageText}\"");

            string chatReply;
            string voiceReply;

            if (lower.Contains("kaise ho") || lower.Contains("kaisa"))
            {
                chatReply = $"Hello {fromUserName}! 🙏 Main badhiya hoon (I'm doing well!). Main TDA Bot hoon, aapka meeting assistant. Kya madad kar sakta hoon?";
                voiceReply = $"Hello {fromUserName}! Main badhiya hoon. I am TDA Bot. How can I help you today?";
            }
            else if (lower.Contains("status") || lower.Contains("recording"))
            {
                chatReply = $"📊 **TDA Bot Status:** \n• 🎙️ Audio recording: Active\n• 📺 Screen share capture: Active\n• 🤖 AI assistant: Listening\n\nMeeting session is being recorded and transcribed.";
                voiceReply = "TDA Bot is active. I am recording meeting audio and capturing screen shares.";
            }
            else if (lower.Contains("help"))
            {
                chatReply = $"Hi {fromUserName}! 👋 I'm **TDA Bot**. I can:\n• 🎙️ Record meeting audio\n• 📺 Capture screen shares\n• 💬 Reply to your messages\n\nJust type **tda bot** or **hi bot** to talk to me!";
                voiceReply = $"Hello {fromUserName}! I am TDA Bot. I am recording the meeting and can assist you.";
            }
            else
            {
                chatReply = $"Hello {fromUserName}! 👋 I'm **TDA Bot**, your AI meeting assistant. I'm actively listening and recording this meeting. Type **tda bot status** to check my status.";
                voiceReply = $"Hello {fromUserName}! I am TDA Bot, your meeting assistant. How can I assist you?";
            }

            _ = Task.Run(async () =>
            {
                await this.PostTextMessageToChatAsync(chatReply).ConfigureAwait(false);
            });

            if (this.AudioSender != null && this.AudioSender.IsAudioSendActive && !this.IsMuted)
            {
                _ = this.AudioSender.SpeakAsync(voiceReply);
            }
        }

        private void OnVoiceTriggerDetected(string triggerText, string voiceReply, string chatReply)
        {
            this.graphLogger?.Info($"[Voice Trigger] Heard: \"{triggerText}\" -> (Muted: {this.IsMuted}) Reply: \"{voiceReply}\"");
            Console.WriteLine($">>> [Voice Trigger] Heard: \"{triggerText}\" -> (Muted: {this.IsMuted})");

            if (this.AudioSender != null && !this.IsMuted)
            {
                _ = this.AudioSender.SpeakAsync(voiceReply);
            }

            if (!string.IsNullOrWhiteSpace(chatReply))
            {
                _ = Task.Run(async () =>
                {
                    await this.PostTextMessageToChatAsync(chatReply).ConfigureAwait(false);
                });
            }
        }

        // ===================================================================
        // 4. Wind-down: video finalisation, audio flush, transcripts, MoM, metadata
        // ===================================================================
        private async Task HandleCallEndedAsync()
        {
            this.graphLogger.Info($"Call {this.Call.Id} terminated - saving all audio, video, transcripts, timeline, MoM and metadata.");

            // 0. Finalise video files first
            this.StopVbssRecorder("call ended");
            this.StopCameraRecorder("call ended");
            this.Timeline.EndedAt = DateTime.Now;

            // 1. Flush Audio to the session folder
            var speakerAudios = this.AudioAggregator.FlushToWavFiles(this.RecordingsManager.SessionDirectory);

            // 2. Transcribe
            var entries = new List<TranscriptEntry>();
            if (speakerAudios.Count > 0)
            {
                foreach (var speakerAudio in speakerAudios)
                {
                    var speakerName = this.ResolveSpeakerName(speakerAudio.SpeakerId);
                    var text = await WhisperTranscriber.TranscribeAsync(speakerAudio.WavPath).ConfigureAwait(false);
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    entries.Add(new TranscriptEntry
                    {
                        Timestamp = speakerAudio.FirstSeenAt.ToString("yyyy-MM-dd HH:mm:ss"),
                        Speaker = speakerName,
                        Transcript = text.Trim(),
                    });
                }
            }

            // 3. Save Transcripts (.txt and .json)
            this.RecordingsManager.SaveTranscripts(entries);

            // 4. Optional MP4 conversion of the recorded video segments
            var allVideoSessions = this.Timeline.ScreenShareSessions.Concat(this.Timeline.CameraSessions).ToList();
            if (!string.IsNullOrWhiteSpace(this.options.FfmpegPath))
            {
                foreach (var session in allVideoSessions)
                {
                    try
                    {
                        session.Mp4Files = await VideoRecorder.ConvertSegmentsToMp4Async(session.VideoFiles, this.options.FfmpegPath, this.RecordingsManager.Log).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        this.graphLogger.Warn($"[Video] MP4 conversion failed: {ex.Message}");
                    }
                }
            }

            // 5. Save timeline (who joined/left, who shared what and when, file names)
            string timelinePath = null;
            try
            {
                timelinePath = this.Timeline.Save(this.RecordingsManager.SessionDirectory);
            }
            catch (Exception ex)
            {
                this.graphLogger.Warn($"[Timeline] Save failed: {ex.Message}");
            }

            // 6. Generate Minutes of Meeting (MoM) - detailed Word doc from transcript + screen snapshots
            if (this.options.Mom?.Enabled == true)
            {
                try
                {
                    var talkTimeByName = speakerAudios.ToDictionary(
                        s => this.ResolveSpeakerName(s.SpeakerId),
                        s => System.IO.File.Exists(s.WavPath) ? Math.Max(0, new FileInfo(s.WavPath).Length - 44) / 32000.0 : 0.0,
                        StringComparer.OrdinalIgnoreCase);

                    var momResult = await TeamsCallingBot.Mom.MomGenerator.GenerateAsync(
                        this.RecordingsManager.SessionDirectory,
                        this.Timeline,
                        entries,
                        talkTimeByName,
                        this.options.Mom,
                        msg => this.RecordingsManager.Log(msg)).ConfigureAwait(false);

                    this.graphLogger.Info($"[MoM] Generated {momResult.DocxPath ?? momResult.MarkdownPath ?? momResult.JsonPath} (AI: {momResult.UsedAi})");
                    Console.WriteLine($">>> [MoM] Generated {(momResult.UsedAi ? "AI-enhanced" : "local")} Minutes of Meeting -> {Path.GetFileName(momResult.DocxPath ?? momResult.MarkdownPath ?? momResult.JsonPath)}");

                    // 6a. Best-effort upload to GCS / Cloud Run relay notification
                    if (!string.IsNullOrWhiteSpace(momResult.DocxPath) && this.options.Gcs != null && !string.IsNullOrWhiteSpace(this.options.Gcs.BucketName))
                    {
                        var gcsUploader = new GcsUploader(this.options.Gcs, msg => this.RecordingsManager.Log(msg));
                        string signedUrl = await gcsUploader.UploadMomDocumentAsync(momResult.DocxPath, this.Call.Id, this.chatThreadId).ConfigureAwait(false);
                        await gcsUploader.UploadTranscriptFilesAsync(this.RecordingsManager.SessionDirectory, this.Call.Id, this.chatThreadId).ConfigureAwait(false);

                        var relayClient = new CloudRunRelayClient(this.options.CloudRunRelay, msg => this.RecordingsManager.Log(msg));
                        if (relayClient.IsConfigured)
                        {
                            await relayClient.SendMomNotificationAsync(
                                momResult.Document,
                                this.chatThreadId,
                                signedUrl,
                                this.options.Gcs?.SignedUrlExpiryHours ?? 168,
                                momResult.UsedAi).ConfigureAwait(false);
                        }
                    }
                }
                catch (Exception ex)
                {
                    this.graphLogger.Error(ex, "[MoM] Generation failed.");
                }
            }

            // 7. Save Session Metadata
            var metadata = new
            {
                SessionId = this.Call.Id,
                StartTime = this.sessionStartTime,
                EndTime = DateTime.Now,
                DurationSeconds = (DateTime.Now - this.sessionStartTime).TotalSeconds,
                SpeakerAudioFilesCount = speakerAudios.Count,
                TranscriptsCount = entries.Count,
                ScreenShareSessions = this.Timeline.ScreenShareSessions.Select(s => new { s.PresenterName, s.StartedAt, s.EndedAt, s.FramesWritten, s.VideoFiles, s.Mp4Files, SnapshotCount = s.SnapshotFiles.Count }),
                CameraSessions = this.Timeline.CameraSessions.Select(s => new { s.PresenterName, s.StartedAt, s.EndedAt, s.FramesWritten, s.VideoFiles, s.Mp4Files }),
                Participants = this.Timeline.Participants.Select(p => new { p.DisplayName, p.IsBot, p.JoinedAt, p.LeftAt }),
                TimelineFile = timelinePath,
                SessionDirectory = this.RecordingsManager.SessionDirectory
            };
            this.RecordingsManager.SaveMetadata(metadata);

            this.graphLogger.Info($"[Session Complete] All assets saved to {this.RecordingsManager.SessionDirectory}");
            Console.WriteLine($">>> [Session Complete] All assets saved to {this.RecordingsManager.SessionDirectory}");
        }

        private string ResolveSpeakerName(uint speakerId)
        {
            if (speakerId == AudioAggregator.UnknownSpeakerId)
            {
                return "Unknown speaker";
            }

            var participant = this.Call.Participants.FirstOrDefault(p =>
                p.Resource?.IsInLobby == false &&
                p.Resource?.MediaStreams != null &&
                p.Resource.MediaStreams.Any(m => m.SourceId == speakerId.ToString()));

            var displayName = participant?.Resource?.Info?.Identity?.User?.DisplayName;
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = this.Timeline.ScreenShareSessions.Concat(this.Timeline.CameraSessions)
                    .FirstOrDefault(s => s.MediaSourceId == speakerId)?.PresenterName;
            }

            return string.IsNullOrWhiteSpace(displayName) ? $"Speaker {speakerId}" : displayName;
        }

        protected override Task HeartbeatAsync(ElapsedEventArgs args)
        {
            return this.Call.KeepAliveAsync();
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            this.videoBroadcastCts?.Cancel();
            this.periodicFlushCts?.Cancel();
            this.autoLeaveCts?.Cancel();
            this.chatMonitorCts?.Cancel();
            this.voiceRecognizer?.Dispose();
            this.AudioSender?.Dispose();

            this.Call.OnUpdated -= this.OnCallUpdated;
            this.Call.Participants.OnUpdated -= this.OnParticipantsUpdated;

            lock (this.participantsWithUpdateHook)
            {
                foreach (var participant in this.Call.Participants)
                {
                    if (this.participantsWithUpdateHook.Contains(participant.Id))
                    {
                        try { participant.OnUpdated -= this.OnParticipantUpdated; } catch { }
                    }
                }

                this.participantsWithUpdateHook.Clear();
            }

            if (this.audioSocket != null)
            {
                this.audioSocket.AudioMediaReceived -= this.OnAudioMediaReceived;
            }

            if (this.vbssSocket != null)
            {
                this.vbssSocket.VideoReceiveStatusChanged -= this.OnVbssReceiveStatusChanged;
                this.vbssSocket.VideoMediaReceived -= this.OnVbssMediaReceived;
                this.vbssSocket.MediaStreamFailure -= this.OnMediaStreamFailure;
            }

            if (this.videoSocket != null)
            {
                this.videoSocket.VideoSendStatusChanged -= this.OnVideoSendStatusChanged;
                this.videoSocket.VideoKeyFrameNeeded -= this.OnVideoKeyFrameNeeded;
                this.videoSocket.VideoReceiveStatusChanged -= this.OnVideoReceiveStatusChanged;
                this.videoSocket.VideoMediaReceived -= this.OnVideoMediaReceived;
                this.videoSocket.MediaStreamFailure -= this.OnMediaStreamFailure;
            }

            this.StopVbssRecorder("handler disposed");
            this.StopCameraRecorder("handler disposed");

            this.chatClient?.Dispose();
            this.latestScreenBitmap?.Dispose();
            lock (this.visualizationLock)
            {
                this.currentVisualization?.Dispose();
                this.currentVisualization = null;
            }
        }
    }
}
