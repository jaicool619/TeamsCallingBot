namespace TeamsCallingBot.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
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
    using TeamsCallingBot.Common;
    using TeamsCallingBot.Config;
    using TeamsCallingBot.Storage;
    using TeamsCallingBot.Video;

    /// <summary>
    /// Complete call lifecycle handler supporting:
    /// 1. Bidirectional Audio (Listen + Speak via AudioSender)
    /// 2. Mute / Unmute
    /// 3. Video Screen Sharing Photo Capture (from participants)
    /// 4. Disk Storage (Audio, Photos, Transcripts, Metadata under D:\Teamsbot\Recordings\<SessionId>)
    /// 5. Bot Video Broadcast (Streaming status card into meeting)
    /// 6. Auto Leave (Leaves after 30s when all human participants leave)
    /// 7. Group Chat Notification if removed by user
    /// </summary>
    public class CallHandler : HeartbeatHandler
    {
        private int endHandled;
        private readonly IGraphLogger graphLogger;
        private readonly DateTime sessionStartTime;

        // Media Sockets
        private readonly IAudioSocket audioSocket;
        private readonly IVideoSocket vbssSocket;
        private readonly IVideoSocket videoSocket;

        // Controllers & Senders
        public AudioSender AudioSender { get; }
        public AudioAggregator AudioAggregator { get; }
        public RecordingsManager RecordingsManager { get; }

        // State Flags
        public bool IsMuted { get; private set; } = false;
        private volatile bool isVideoSendActive = false;
        private volatile bool isKeyFrameNeeded = true;
        private int joinWelcomeSent = 0;

        // Screen & Video Capture Throttling (independent throttles so screen capture isn't starved)
        private DateTime lastVbssPhotoTime = DateTime.MinValue;
        private DateTime lastVideoPhotoTime = DateTime.MinValue;
        private readonly TimeSpan photoInterval = TimeSpan.FromSeconds(5);
        private readonly object photoLock = new object();
        private Bitmap latestScreenBitmap = null;

        // Meeting Chat Polling & Dynamic Conversational Reply Loop
        private CancellationTokenSource chatMonitorCts;
        private readonly HashSet<string> processedChatMsgIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Bot Video Streaming
        private CancellationTokenSource videoBroadcastCts;

        // Periodic Audio Flush
        private CancellationTokenSource periodicFlushCts;

        // Auto-Leave Timer
        private CancellationTokenSource autoLeaveCts;
        private const int AutoLeaveDebounceSeconds = 120;
        private const int InitialGracePeriodSeconds = 300;

        // Chat Info
        private readonly string chatThreadId;
        private readonly string botAccessToken;

        // Real-Time Voice Speech Recognition
        private readonly RealtimeSpeechRecognizer voiceRecognizer;

        public ICall Call { get; }

        public CallHandler(ICall call, IGraphLogger logger, string chatThreadId = null, string accessToken = null)
            : base(TimeSpan.FromMinutes(1), logger)
        {
            this.Call = call ?? throw new ArgumentNullException(nameof(call));
            this.graphLogger = logger;
            this.sessionStartTime = DateTime.Now;
            this.chatThreadId = chatThreadId;
            this.botAccessToken = accessToken ?? BotOptions.Current?.OverrideBearerToken;

            // 1. Initialize Disk Storage Manager
            this.RecordingsManager = new RecordingsManager(this.Call.Id);
            this.AudioAggregator = new AudioAggregator();

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

                    // Play pleasant greeting chime + verbal announcement when audio send is active
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
                        }
                    });
                }

                // VBSS (Screen Share Receive) Socket
                this.vbssSocket = localMediaSession.VbssSocket;
                if (this.vbssSocket != null)
                {
                    this.vbssSocket.VideoReceiveStatusChanged += this.OnVbssReceiveStatusChanged;
                    this.vbssSocket.VideoMediaReceived += this.OnVbssMediaReceived;

                    try
                    {
                        this.vbssSocket.Subscribe(VideoResolution.HD1080p);
                    }
                    catch (Exception ex)
                    {
                        this.graphLogger.Warn($"Initial VBSS subscribe returned: {ex.Message}");
                    }

                    this.graphLogger.Info($"[VBSS Socket Initialized] Screen sharing photo capture is active.");
                }

                // Video Socket (Bot Video Streaming & Meeting Video Receiving)
                if (localMediaSession.VideoSockets != null && localMediaSession.VideoSockets.Any())
                {
                    this.videoSocket = localMediaSession.VideoSockets.First();
                    this.videoSocket.VideoSendStatusChanged += this.OnVideoSendStatusChanged;
                    this.videoSocket.VideoKeyFrameNeeded += this.OnVideoKeyFrameNeeded;
                    this.videoSocket.VideoReceiveStatusChanged += this.OnVideoReceiveStatusChanged;
                    this.videoSocket.VideoMediaReceived += this.OnVideoMediaReceived;
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

            // 5. Start Active Meeting Chat Monitor (replies to "tda bot", "hi", "kaise ho" in group chat + voice)
            this.StartMeetingChatMonitor();

            // 6. Start Real-Time Voice Speech Recognizer (replies to spoken voice: "tda bot", "kaise ho", etc.)
            this.voiceRecognizer = new RealtimeSpeechRecognizer(
                this.graphLogger,
                () => this.AudioSender != null && this.AudioSender.IsSpeaking);
            this.voiceRecognizer.OnTriggerDetected += this.OnVoiceTriggerDetected;

            // Diagnostic: log what auth credentials are available for chat posting
            string diagAppId = BotOptions.Current?.AadAppId ?? "(null)";
            string diagSecret = string.IsNullOrWhiteSpace(BotOptions.Current?.AadAppSecretOrCertThumbprint) ? "(empty)" : $"({BotOptions.Current.AadAppSecretOrCertThumbprint.Length} chars)";
            string diagOverride = string.IsNullOrWhiteSpace(BotOptions.Current?.OverrideBearerToken) ? "(empty)" : $"({BotOptions.Current.OverrideBearerToken.Length} chars)";
            Console.WriteLine($">>> [Config Diagnostic] AppId={diagAppId}, Secret={diagSecret}, OverrideToken={diagOverride}");

            this.graphLogger.Info($"CallHandler initialized for {this.Call.Id}. Output folder: {this.RecordingsManager.SessionDirectory}");
            Console.WriteLine($">>> [CallHandler] Output folder: {this.RecordingsManager.SessionDirectory}");
        }

        // ===================================================================
        // 1. Audio (Both Ways)
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
        // 3. Video Screen Sharing & Meeting Video Capture (Photos)
        // ===================================================================
        private void OnVbssReceiveStatusChanged(object sender, VideoReceiveStatusChangedEventArgs e)
        {
            this.graphLogger.Info($"[VBSS Socket] VideoReceiveStatusChanged: {e.MediaReceiveStatus}");
            Console.WriteLine($">>> [VBSS Socket] VideoReceiveStatus: {e.MediaReceiveStatus}");
            if (e.MediaReceiveStatus == MediaReceiveStatus.Active && this.vbssSocket != null)
            {
                try
                {
                    this.vbssSocket.Subscribe(VideoResolution.HD1080p);
                    this.graphLogger.Info("[VBSS Socket] Subscribed to HD1080p screen share.");
                    Console.WriteLine(">>> [VBSS Socket] Subscribed to HD1080p screen share.");
                }
                catch (Exception ex)
                {
                    this.graphLogger.Warn($"[VBSS Socket] Subscribe returned: {ex.Message}");
                }
            }
        }

        private void OnVbssMediaReceived(object sender, VideoMediaReceivedEventArgs e)
        {
            try
            {
                DateTime now = DateTime.Now;
                bool shouldCapture = false;

                lock (this.photoLock)
                {
                    if (now - this.lastVbssPhotoTime >= this.photoInterval)
                    {
                        this.lastVbssPhotoTime = now;
                        shouldCapture = true;
                    }
                }

                if (shouldCapture && e.Buffer.Data != IntPtr.Zero)
                {
                    Bitmap bitmap = null;
                    if (e.Buffer.VideoFormat?.VideoColorFormat == VideoColorFormat.NV12)
                    {
                        bitmap = VideoFrameConverter.ConvertNV12ToBitmap(e.Buffer.Data, e.Buffer.VideoFormat.Width, e.Buffer.VideoFormat.Height, e.Buffer.Stride);
                    }
                    else if (e.Buffer.VideoFormat?.VideoColorFormat == VideoColorFormat.Rgb24)
                    {
                        bitmap = VideoFrameConverter.ConvertRGB24ToBitmap(e.Buffer.Data, e.Buffer.VideoFormat.Width, e.Buffer.VideoFormat.Height, e.Buffer.Stride);
                    }

                    if (bitmap != null)
                    {
                        lock (this.photoLock)
                        {
                            this.latestScreenBitmap?.Dispose();
                            this.latestScreenBitmap = (Bitmap)bitmap.Clone();
                        }

                        var savedPath = this.RecordingsManager.SavePhoto(bitmap, "03_photo_screenshare");
                        this.graphLogger.Info($"[Screen Photo Saved] {savedPath}");
                        Console.WriteLine($">>> [Screen Photo Saved] {savedPath}");
                        bitmap.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                this.graphLogger.Error(ex, "Error capturing screen share frame.");
            }
            finally
            {
                e.Buffer.Dispose();
            }
        }

        private void OnVideoReceiveStatusChanged(object sender, VideoReceiveStatusChangedEventArgs e)
        {
            this.graphLogger.Info($"[Video Socket] VideoReceiveStatusChanged: {e.MediaReceiveStatus}");
            Console.WriteLine($">>> [Video Socket] VideoReceiveStatus: {e.MediaReceiveStatus}");
            if (e.MediaReceiveStatus == MediaReceiveStatus.Active && this.videoSocket != null)
            {
                try
                {
                    this.videoSocket.Subscribe(VideoResolution.HD720p);
                    this.graphLogger.Info("[Video Socket] Subscribed to HD720p meeting video.");
                    Console.WriteLine(">>> [Video Socket] Subscribed to HD720p meeting video.");
                }
                catch (Exception ex)
                {
                    this.graphLogger.Warn($"[Video Socket] Subscribe returned: {ex.Message}");
                }
            }
        }

        private void OnVideoMediaReceived(object sender, VideoMediaReceivedEventArgs e)
        {
            try
            {
                DateTime now = DateTime.Now;
                bool shouldCapture = false;

                lock (this.photoLock)
                {
                    if (now - this.lastVideoPhotoTime >= this.photoInterval)
                    {
                        this.lastVideoPhotoTime = now;
                        shouldCapture = true;
                    }
                }

                if (shouldCapture && e.Buffer.Data != IntPtr.Zero)
                {
                    Bitmap bitmap = null;
                    if (e.Buffer.VideoFormat?.VideoColorFormat == VideoColorFormat.NV12)
                    {
                        bitmap = VideoFrameConverter.ConvertNV12ToBitmap(e.Buffer.Data, e.Buffer.VideoFormat.Width, e.Buffer.VideoFormat.Height, e.Buffer.Stride);
                    }
                    else if (e.Buffer.VideoFormat?.VideoColorFormat == VideoColorFormat.Rgb24)
                    {
                        bitmap = VideoFrameConverter.ConvertRGB24ToBitmap(e.Buffer.Data, e.Buffer.VideoFormat.Width, e.Buffer.VideoFormat.Height, e.Buffer.Stride);
                    }

                    if (bitmap != null)
                    {
                        lock (this.photoLock)
                        {
                            this.latestScreenBitmap?.Dispose();
                            this.latestScreenBitmap = (Bitmap)bitmap.Clone();
                        }

                        var savedPath = this.RecordingsManager.SavePhoto(bitmap, "03_photo_video");
                        this.graphLogger.Info($"[Meeting Video Photo Saved] {savedPath}");
                        Console.WriteLine($">>> [Meeting Video Photo Saved] {savedPath}");
                        bitmap.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                this.graphLogger.Error(ex, "Error capturing meeting video frame.");
            }
            finally
            {
                e.Buffer.Dispose();
            }
        }

        /// <summary>
        /// Instantly saves a photo of the current active screen share or meeting video on demand.
        /// </summary>
        public string CapturePhotoNow(string tag = "manual")
        {
            lock (this.photoLock)
            {
                if (this.latestScreenBitmap != null)
                {
                    return this.RecordingsManager.SavePhoto(this.latestScreenBitmap, $"03_photo_{tag}");
                }
            }
            return null;
        }

        // ===================================================================
        // 5. Video Sharing by Bot
        // ===================================================================
        private void OnVideoSendStatusChanged(object sender, VideoSendStatusChangedEventArgs e)
        {
            this.isVideoSendActive = (e.MediaSendStatus == MediaSendStatus.Active);
            this.graphLogger.Info($"[VideoSocket] VideoSendStatusChanged: {e.MediaSendStatus}");
            Console.WriteLine($">>> [VideoSocket] VideoSendStatus: {e.MediaSendStatus}");
        }

        private void OnVideoKeyFrameNeeded(object sender, VideoKeyFrameNeededEventArgs e)
        {
            this.isKeyFrameNeeded = true;
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
                int frameRate = 15; // 15 FPS
                int frameDelayMs = 1000 / frameRate;
                uint timestamp = 0;
                int tick = 0;
                bool snapshotSaved = false;
                bool loggedFirstSend = false;
                bool loggedFirstError = false;
                int consecutiveErrors = 0;
                int bufSize = width * height * 3 / 2; // 1,382,400 bytes for NV12 1280x720

                // CRITICAL STABILITY FIX: Pre-allocate a pool of 6 unmanaged video frame buffers (8.3 MB total).
                // Standard VideoSendBuffer calls Marshal.FreeHGlobal() immediately upon Dispose(), which causes
                // native video encoder worker threads in mpencoder.dll / MediaApi.dll to access freed memory
                // and crash the process with Access Violation (0xc0000005).
                // With SafeVideoMediaBuffer and a ring buffer pool, each slot has 6 * 66ms = 400ms safety
                // window for the native encoder to finish reading the frame before the slot is reused.
                const int poolSize = 6;
                IntPtr[] videoBufferPool = new IntPtr[poolSize];
                for (int i = 0; i < poolSize; i++)
                {
                    videoBufferPool[i] = Marshal.AllocHGlobal(bufSize);
                }

                // Wait for call to be established before sending frames
                await Task.Delay(2000, token).ConfigureAwait(false);

                try
                {
                    while (!token.IsCancellationRequested)
                    {
                        tick++;

                        // Generate dynamic status card with prominent bot icon, waveform bars, and status
                        using (var card = VideoFrameConverter.CreateBotStatusCard(
                            "Teams AI Assistant",
                            this.IsMuted ? "Muted" : "Active & Listening",
                            this.Call.Id,
                            this.IsMuted,
                            tick))
                        {
                            // Save a single preview snapshot of the bot broadcast card on start
                            if (!snapshotSaved)
                            {
                                snapshotSaved = true;
                                this.RecordingsManager.SavePhoto(card, "05_bot_broadcast_preview");
                            }

                            if (this.isKeyFrameNeeded)
                            {
                                this.isKeyFrameNeeded = false;
                            }

                            byte[] nv12Bytes = VideoFrameConverter.ConvertBitmapToNV12(card, width, height);

                            try
                            {
                                int slot = tick % poolSize;
                                Marshal.Copy(nv12Bytes, 0, videoBufferPool[slot], bufSize);

                                var videoBuffer = new SafeVideoMediaBuffer(
                                    videoBufferPool[slot],
                                    bufSize,
                                    VideoFormat.NV12_1280x720_30Fps,
                                    (long)timestamp);

                                this.videoSocket.Send(videoBuffer);

                                consecutiveErrors = 0;
                                if (!loggedFirstSend)
                                {
                                    loggedFirstSend = true;
                                    Console.WriteLine($">>> [Bot Video] First frame sent successfully at tick {tick}! Bot should now be visible.");
                                    this.graphLogger.Info($"[Bot Video] First frame sent successfully at tick {tick}.");
                                }
                            }
                            catch (Exception sendEx)
                            {
                                consecutiveErrors++;
                                if (!loggedFirstError)
                                {
                                    loggedFirstError = true;
                                    Console.WriteLine($">>> [Bot Video] Send attempt failed (will keep retrying): {sendEx.GetType().Name}: {sendEx.Message}");
                                }
                                // After initial errors, slow down to avoid spamming
                                if (consecutiveErrors > 30 && consecutiveErrors % 100 == 0)
                                {
                                    Console.WriteLine($">>> [Bot Video] Still retrying... ({consecutiveErrors} consecutive send failures, isVideoSendActive={this.isVideoSendActive})");
                                }
                            }
                        }

                        timestamp += (uint)frameDelayMs;
                        await Task.Delay(frameDelayMs, token).ConfigureAwait(false);
                    }
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    this.graphLogger.Error(ex, "Error in bot video broadcast loop.");
                    Console.WriteLine($">>> [Bot Video] FATAL broadcast loop error: {ex.Message}");
                }
                finally
                {
                    // Clean up unmanaged video buffers when broadcast loop exits
                    for (int i = 0; i < poolSize; i++)
                    {
                        if (videoBufferPool[i] != IntPtr.Zero)
                        {
                            Marshal.FreeHGlobal(videoBufferPool[i]);
                            videoBufferPool[i] = IntPtr.Zero;
                        }
                    }
                }
            }, token);

            this.graphLogger.Info("[Bot Video Streaming] Broadcast loop initialized at 15 FPS.");
        }

        // ===================================================================
        // 6. Auto Leave
        // ===================================================================
        private void OnParticipantsUpdated(IParticipantCollection sender, CollectionEventArgs<IParticipant> args)
        {
            string botAppId = BotOptions.Current?.AadAppId;

            // Count participants who are in the meeting (not in lobby, and not this bot)
            int humanCount = this.Call.Participants.Count(p =>
                p.Resource.IsInLobby == false &&
                (p.Resource.Info?.Identity?.Application == null || p.Resource.Info.Identity.Application.Id != botAppId));

            this.graphLogger.Info($"[Participants Updated] Active non-bot participant count: {humanCount}");
            Console.WriteLine($">>> [Participants Updated] Active non-bot participant count: {humanCount}");

            // Check and subscribe to participant screen share and video streams
            this.SubscribeParticipantMediaStreams();

            // Initial Grace Period check: do not auto-leave during the first 45 seconds of the call
            bool inGracePeriod = (DateTime.Now - this.sessionStartTime).TotalSeconds < InitialGracePeriodSeconds;

            if (humanCount == 0 && !inGracePeriod)
            {
                // All non-bot participants left after the initial grace period - start countdown
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
                // Humans present - cancel any pending auto leave
                if (this.autoLeaveCts != null && !this.autoLeaveCts.IsCancellationRequested)
                {
                    this.autoLeaveCts.Cancel();
                    this.autoLeaveCts = null;
                    this.graphLogger.Info("[Auto Leave] Canceled because a human participant is present.");
                    Console.WriteLine(">>> [Auto Leave] Countdown canceled - participant is present.");
                }
            }
        }

        private void SubscribeParticipantMediaStreams()
        {
            try
            {
                string botAppId = BotOptions.Current?.AadAppId;
                if (this.Call?.Participants == null) return;

                foreach (var participant in this.Call.Participants)
                {
                    if (participant.Resource.IsInLobby == true) continue;
                    if (participant.Resource.Info?.Identity?.Application?.Id == botAppId) continue;

                    var streams = participant.Resource.MediaStreams;
                    if (streams == null) continue;

                    foreach (var stream in streams)
                    {
                        if (uint.TryParse(stream.SourceId, out uint msi) && msi > 0)
                        {
                            string label = stream.Label ?? "";
                            string mediaTypeStr = stream.MediaType.ToString();

                            // CRITICAL FIX: Only subscribe to a participant MSI if the stream is actively sending!
                            // If stream.Direction is Inactive or ReceiveOnly, subscribing to it will overwrite and break the active stream!
                            bool isSending = stream.Direction == MediaDirection.SendOnly || stream.Direction == MediaDirection.SendReceive;
                            if (!isSending)
                            {
                                continue;
                            }

                            // 1. Screen sharing stream (VBSS)
                            if (mediaTypeStr.IndexOf("ScreenSharing", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                label.IndexOf("sharing", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                try
                                {
                                    this.vbssSocket?.Subscribe(VideoResolution.HD1080p, msi);
                                    this.graphLogger.Info($"[VBSS Socket] Subscribed to active screen share MSI {msi} ({label}) for {participant.Id}");
                                    Console.WriteLine($">>> [VBSS Socket] Subscribed to active screen share MSI {msi} ({label})");
                                }
                                catch (Exception ex)
                                {
                                    this.graphLogger.Warn($"[VBSS Socket] Error subscribing MSI {msi}: {ex.Message}");
                                }
                            }
                            // 2. Participant camera / video stream
                            else if (mediaTypeStr.IndexOf("Video", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                     label.IndexOf("video", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                try
                                {
                                    this.videoSocket?.Subscribe(VideoResolution.HD720p, msi);
                                    this.graphLogger.Info($"[Video Socket] Subscribed to meeting video MSI {msi} ({label}) for {participant.Id}");
                                    Console.WriteLine($">>> [Video Socket] Subscribed to meeting video MSI {msi} ({label})");
                                }
                                catch (Exception ex)
                                {
                                    this.graphLogger.Warn($"[Video Socket] Error subscribing MSI {msi}: {ex.Message}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.graphLogger.Warn($"Error subscribing participant media streams: {ex.Message}");
            }
        }

        // ===================================================================
        // 7. Detection of Bot Removal & Group Chat Message Posting
        // ===================================================================
        private void OnCallUpdated(ICall sender, ResourceEventArgs<Call> args)
        {
            Console.WriteLine($">>> CALL STATE CHANGED: {args.NewResource.State} (call {this.Call.Id})");
            var resultInfo = args.NewResource.ResultInfo;
            if (resultInfo != null)
            {
                Console.WriteLine($">>> CALL RESULT INFO: code={resultInfo.Code}, subcode={resultInfo.Subcode}, message={resultInfo.Message}");
            }

            // Post welcome message and subscribe streams when call becomes Established
            if (args.NewResource.State == CallState.Established)
            {
                if (Interlocked.Exchange(ref this.joinWelcomeSent, 1) == 0)
                {
                    _ = Task.Run(() => this.PostTextMessageToChatAsync(
                        "🤖 <b>Teams AI Assistant</b> has joined the meeting.<br/>• Audio recording & unmixed speaker capture: <b>Active</b><br/>• Video status broadcast: <b>Active</b><br/>• Screen share photo capture: <b>Ready</b>"));
                }

                try
                {
                    this.vbssSocket?.Subscribe(VideoResolution.HD1080p);
                    this.videoSocket?.Subscribe(VideoResolution.HD720p);
                }
                catch { }

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

            // Check if removed by user / organizer (subcodes 7002, 702, 5000, 403, or keyword messages)
            bool wasRemovedByUser = false;
            if (resultInfo != null)
            {
                string msg = resultInfo.Message?.ToLowerInvariant() ?? "";
                if (resultInfo.Subcode == 7002 || resultInfo.Subcode == 702 || resultInfo.Subcode == 5000 || resultInfo.Code == 403 ||
                    msg.Contains("removed") || msg.Contains("kicked") || msg.Contains("ejected") || msg.Contains("organizer"))
                {
                    wasRemovedByUser = true;
                }
            }

            if (wasRemovedByUser)
            {
                Console.WriteLine(">>> BOT REMOVAL DETECTED: Announcing verbally and posting notification to group chat...");
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
                    "👋 <b>Teams Calling Bot Notification:</b> The bot has left the meeting. All session audio, screen captures, and transcripts have been archived.",
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

        /// <summary>
        /// Acquires a Bot Framework OAuth token using app credentials (audience: https://api.botframework.com).
        /// This bypasses Microsoft Graph's ChatMessage.Send permission requirements completely!
        /// </summary>
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

            // Endpoints to attempt: specific tenant endpoint (do not use /common/ for single-tenant apps)
            var tokenEndpoints = !string.IsNullOrWhiteSpace(tenantId) && !tenantId.Equals("common", StringComparison.OrdinalIgnoreCase)
                ? new[] { $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token" }
                : new[] { "https://login.microsoftonline.com/common/oauth2/v2.0/token" };

            foreach (var tokenEndpoint in tokenEndpoints)
            {
                try
                {
                    Console.WriteLine($">>> [BF Token] Requesting token from {tokenEndpoint} for api.botframework.com...");
                    logger?.Info($"[BF Token] Requesting token from {tokenEndpoint} for api.botframework.com...");

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

                            Console.WriteLine($">>> [BF Token] SUCCESS - token acquired from {tokenEndpoint}, expires in {expiresIn}s");
                            logger?.Info($"[BF Token] SUCCESS - token acquired from {tokenEndpoint}, expires in {expiresIn}s");
                            return token;
                        }
                        else
                        {
                            Console.WriteLine($">>> [BF Token] {tokenEndpoint} FAILED ({(int)resp.StatusCode}): {json}");
                            logger?.Warn($"[BF Token] {tokenEndpoint} FAILED ({(int)resp.StatusCode}): {json}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($">>> [BF Token] EXCEPTION on {tokenEndpoint}: {ex.GetType().Name}: {ex.Message}");
                    logger?.Error(ex, $"[BF Token] EXCEPTION on {tokenEndpoint}");
                }
            }

            return null;
        }

        /// <summary>
        /// Posts an activity to a Teams conversation via the official Bot Framework Connector API.
        /// Standard for Teams bots, requires NO Microsoft Graph ChatMessage.Send.* permissions.
        /// The activity MUST include from, conversation, channelId, and serviceUrl fields or Teams will reject it.
        /// </summary>
        public async Task<bool> PostActivityViaBotFrameworkAsync(string threadId, string text)
        {
            string bfToken = await GetBotFrameworkTokenAsync(this.graphLogger).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(bfToken))
            {
                this.graphLogger.Warn("[BotFramework Connector] Could not acquire Bot Framework token.");
                return false;
            }

            string botAppId = BotOptions.Current?.AadAppId ?? string.Empty;
            string tenantId = BotOptions.Current?.AadTenantId ?? string.Empty;

            // Try multiple SMBA regional endpoints.
            // IMPORTANT: Do NOT Uri.EscapeDataString the threadId - it must be passed verbatim
            // (Teams thread IDs like "19:meeting_xxx@thread.v2" are already valid path segments).
            var serviceUrls = new[]
            {
                "https://smba.trafficmanager.net/apis/",
                "https://smba.trafficmanager.net/apac/",
                "https://smba.trafficmanager.net/amer/",
                "https://smba.trafficmanager.net/emea/"
            };

            // Full Bot Framework activity - 'from', 'conversation', 'channelId', 'serviceUrl' are required.
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
                        else
                        {
                            string errBody = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                            this.graphLogger.Warn($"[BotFramework Connector] {serviceUrl} returned {(int)resp.StatusCode}: {errBody}");
                            Console.WriteLine($">>> [BotFramework Connector] {serviceUrl} returned {(int)resp.StatusCode}: {errBody}");

                            if (errBody != null && errBody.IndexOf("BotNotInConversationRoster", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                this.graphLogger.Warn("[BotFramework Connector] BotNotInConversationRoster: The bot app is not yet added to this meeting's roster. To enable meeting chat posting, a participant should click '+' (Apps) in Teams meeting and add the bot app.");
                                Console.WriteLine(">>> [BotFramework Connector] HINT: Add the bot app to this meeting via '+' (Apps) in Teams to allow it in the meeting chat roster.");
                                return false;
                            }

                            // If 401 on first attempt, the token is invalid — bail immediately.
                            if (resp.StatusCode == System.Net.HttpStatusCode.Unauthorized) return false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    this.graphLogger.Warn($"[BotFramework Connector] Exception on {serviceUrl}: {ex.Message}");
                }
            }

            return false;
        }

        /// <summary>
        /// Posts a formatted text/HTML message into the Teams meeting group chat.
        /// Tries Bot Framework Connector API first (no Graph permission needed), then falls back to Graph.
        /// </summary>
        public async Task<bool> PostTextMessageToChatAsync(string htmlContent, bool isRemovalNotification = false)
        {
            string threadId = this.GetEffectiveChatThreadId();
            if (string.IsNullOrWhiteSpace(threadId))
            {
                this.graphLogger.Warn("[Chat Notification] Missing chatThreadId - cannot post message.");
                return false;
            }

            // 1. PRIMARY: Bot Framework Connector API (official bot messaging, no Graph admin consent needed)
            string markdownText = htmlContent
                .Replace("<b>", "**").Replace("</b>", "**")
                .Replace("<i>", "*").Replace("</i>", "*");

            bool bfSuccess = await this.PostActivityViaBotFrameworkAsync(threadId, markdownText).ConfigureAwait(false);
            if (bfSuccess)
            {
                if (isRemovalNotification)
                {
                    this.RecordingsManager.SaveRemovalMessageRecord(threadId, htmlContent, true);
                }
                return true;
            }

            // 2. FALLBACK: Microsoft Graph Chat API
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
                        return true;
                    }
                    else
                    {
                        var err = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        this.graphLogger.Warn($"[Chat Notification] Graph chat post returned ({response.StatusCode}): {err}");
                        Console.WriteLine($">>> [Chat Notification] Graph chat post returned ({response.StatusCode}): {err}");
                        if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                        {
                            Console.WriteLine($">>> [ACTION NEEDED] Azure AD missing Application permission 'ChatMessage.Send.Chat' or 'Chat.ReadWrite.All'. Please grant Admin Consent in Azure Portal.");
                        }
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

        /// <summary>
        /// Initializes the CTS for chat monitoring. Actual incoming chat messages are received
        /// via the Bot Framework /api/messages endpoint (BotMessagingController) and routed to
        /// HandleIncomingChatActivity() below. This avoids needing Chat.Read.All Graph permissions.
        /// </summary>
        private void StartMeetingChatMonitor()
        {
            this.chatMonitorCts = new CancellationTokenSource();
            string threadId = this.GetEffectiveChatThreadId();
            Console.WriteLine($">>> [Chat Monitor] Initialized. Meeting chat thread: {threadId ?? "(not yet known - will be set from call.Resource.ChatInfo)"}. Listening via /api/messages.");
        }

        /// <summary>
        /// Called by BotMessagingController when Teams delivers an incoming chat message activity
        /// to /api/messages. Replies both in chat and via voice.
        /// </summary>
        public void HandleIncomingChatActivity(string fromUserName, string fromUserId, string messageText, string incomingServiceUrl)
        {
            if (string.IsNullOrWhiteSpace(messageText)) return;

            string lower = messageText.ToLowerInvariant();

            // Only respond to messages that mention/address the bot
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

            // If we got a serviceUrl from the incoming message, use it for the reply (most reliable)
            string effectiveThreadId = this.GetEffectiveChatThreadId();
            if (!string.IsNullOrWhiteSpace(effectiveThreadId))
            {
                _ = Task.Run(async () =>
                {
                    bool sent = await this.PostActivityViaBotFrameworkAsync(effectiveThreadId, chatReply).ConfigureAwait(false);
                    if (!sent)
                    {
                        // Fallback to Graph if BF fails
                        await this.PostTextMessageToChatAsync(chatReply).ConfigureAwait(false);
                    }
                });
            }

            // Speak response aloud into meeting
            if (this.AudioSender != null && this.AudioSender.IsAudioSendActive)
            {
                _ = this.AudioSender.SpeakAsync(voiceReply);
            }
        }

        /// <summary>
        /// Handles real-time speech recognition triggers spoken aloud into the meeting microphone
        /// (such as "tda bot", "kaise ho", "hi bot", "status", "help"). Replies immediately via voice and chat.
        /// </summary>
        private void OnVoiceTriggerDetected(string triggerText, string voiceReply, string chatReply)
        {
            this.graphLogger?.Info($"[Voice Trigger] Heard: \"{triggerText}\" -> Speaking aloud: \"{voiceReply}\"");
            Console.WriteLine($">>> [Voice Trigger] Heard: \"{triggerText}\" -> Speaking aloud: \"{voiceReply}\"");

            // 1. Speak aloud into the meeting audio stream immediately
            if (this.AudioSender != null)
            {
                _ = this.AudioSender.SpeakAsync(voiceReply);
            }

            // 2. Also attempt to post the reply to the meeting chat
            string effectiveThreadId = this.GetEffectiveChatThreadId();
            if (!string.IsNullOrWhiteSpace(effectiveThreadId) && !string.IsNullOrWhiteSpace(chatReply))
            {
                _ = Task.Run(async () =>
                {
                    bool sent = await this.PostActivityViaBotFrameworkAsync(effectiveThreadId, chatReply).ConfigureAwait(false);
                    if (!sent)
                    {
                        await this.PostTextMessageToChatAsync(chatReply).ConfigureAwait(false);
                    }
                });
            }
        }

        // ===================================================================
        // 4. Transcript & Audio Disk Saving (Wind-Down)
        // ===================================================================
        private async Task HandleCallEndedAsync()
        {
            this.graphLogger.Info($"Call {this.Call.Id} terminated - saving all audio, transcripts and metadata.");

            // 1. Flush Audio to D:\Teamsbot\Recordings\<SessionId>\ (Same folder)
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

            // 4. Save Session Metadata
            var metadata = new
            {
                SessionId = this.Call.Id,
                StartTime = this.sessionStartTime,
                EndTime = DateTime.Now,
                DurationSeconds = (DateTime.Now - this.sessionStartTime).TotalSeconds,
                SpeakerAudioFilesCount = speakerAudios.Count,
                TranscriptsCount = entries.Count,
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

            var participant = this.Call.Participants.SingleOrDefault(p =>
                p.Resource.IsInLobby == false &&
                p.Resource.MediaStreams.Any(m => m.SourceId == speakerId.ToString()));

            var displayName = participant?.Resource?.Info?.Identity?.User?.DisplayName;
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

            if (this.audioSocket != null)
            {
                this.audioSocket.AudioMediaReceived -= this.OnAudioMediaReceived;
            }

            if (this.vbssSocket != null)
            {
                this.vbssSocket.VideoReceiveStatusChanged -= this.OnVbssReceiveStatusChanged;
                this.vbssSocket.VideoMediaReceived -= this.OnVbssMediaReceived;
            }

            if (this.videoSocket != null)
            {
                this.videoSocket.VideoSendStatusChanged -= this.OnVideoSendStatusChanged;
                this.videoSocket.VideoKeyFrameNeeded -= this.OnVideoKeyFrameNeeded;
            }

            this.latestScreenBitmap?.Dispose();
        }
    }
}
