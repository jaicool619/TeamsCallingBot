namespace TeamsCallingBot.Audio
{
    using System;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Graph.Communications.Calls.Media;
    using Microsoft.Graph.Communications.Common.Telemetry;
    using Microsoft.Skype.Bots.Media;

    /// <summary>
    /// Handles outbound audio playback into the Teams meeting through AudioSocket.Send().
    /// Supports sending 16kHz 16-bit Mono PCM audio in 20ms chunks (640 bytes per frame).
    /// </summary>
    public class AudioSender : IDisposable
    {
        private const int SampleRate = 16000;
        private const int BytesPerSample = 2; // 16-bit
        private const int Channels = 1;       // Mono
        private const int FrameDurationMs = 20;
        private const int BytesPerFrame = (SampleRate * BytesPerSample * Channels * FrameDurationMs) / 1000; // 640 bytes

        private readonly IAudioSocket audioSocket;
        private readonly IGraphLogger graphLogger;
        private CancellationTokenSource currentPlaybackCts;
        private readonly object playbackLock = new object();

        public bool IsMuted { get; set; } = false;
        public bool IsAudioSendActive { get; private set; } = false;
        public bool IsSpeaking { get; private set; } = false;

        public AudioSender(IAudioSocket audioSocket, IGraphLogger graphLogger)
        {
            this.audioSocket = audioSocket;
            this.graphLogger = graphLogger;

            if (this.audioSocket != null)
            {
                this.audioSocket.AudioSendStatusChanged += this.OnAudioSendStatusChanged;

                if (this.audioSocket != null)
                {
                    try
                    {
                        var prop = this.audioSocket.GetType().GetProperty("LastKnownSendStatus", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                        if (prop != null)
                        {
                            var val = (MediaSendStatus)prop.GetValue(this.audioSocket);
                            this.IsAudioSendActive = (val == MediaSendStatus.Active);
                        }
                        else
                        {
                            this.IsAudioSendActive = true;
                        }
                    }
                    catch
                    {
                        this.IsAudioSendActive = true;
                    }

                    this.graphLogger?.Info($"[AudioSender] Initialized. IsAudioSendActive={this.IsAudioSendActive}");
                }
            }
        }

        private void OnAudioSendStatusChanged(object sender, AudioSendStatusChangedEventArgs e)
        {
            this.IsAudioSendActive = (e.MediaSendStatus == MediaSendStatus.Active);
            this.graphLogger?.Info($"[AudioSender] AudioSendStatusChanged: {e.MediaSendStatus}");
            Console.WriteLine($">>> [AudioSender] AudioSendStatus: {e.MediaSendStatus}");
        }

        /// <summary>
        /// Plays a standard WAV file (converted/extracted to 16kHz 16-bit mono PCM) into the meeting.
        /// </summary>
        public async Task PlayWavFileAsync(string wavFilePath, CancellationToken cancellationToken = default)
        {
            if (this.audioSocket == null || !File.Exists(wavFilePath))
            {
                this.graphLogger?.Warn($"AudioSender: WAV file not found: {wavFilePath}");
                return;
            }

            var pcmBytes = ExtractPcmFromWav(wavFilePath);
            if (pcmBytes == null || pcmBytes.Length == 0)
            {
                this.graphLogger?.Warn($"AudioSender: Failed to extract PCM bytes from {wavFilePath}");
                return;
            }

            await PlayPcmBytesAsync(pcmBytes, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Plays a sequence of raw 16kHz 16-bit mono PCM bytes in 20ms frames.
        /// Safely allocates per frame so AudioSendBuffer.Dispose() can free its memory without double-freeing.
        /// </summary>
        public async Task PlayPcmBytesAsync(byte[] pcmData, CancellationToken cancellationToken = default)
        {
            if (this.audioSocket == null || pcmData == null || pcmData.Length == 0)
            {
                return;
            }

            lock (this.playbackLock)
            {
                this.currentPlaybackCts?.Cancel();
                this.currentPlaybackCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            }

            var token = this.currentPlaybackCts.Token;

            // Wait up to 5 seconds if audio send is not yet active
            var startWait = DateTime.Now;
            while (!this.IsAudioSendActive && (DateTime.Now - startWait).TotalSeconds < 5 && !token.IsCancellationRequested)
            {
                await Task.Delay(200, token).ConfigureAwait(false);
            }

            uint timestamp = 0;
            int offset = 0;
            var frameBuffer = new byte[BytesPerFrame];
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            long targetElapsedMs = 0;

            const int poolSize = 16;
            IntPtr[] audioPool = new IntPtr[poolSize];
            for (int i = 0; i < poolSize; i++)
            {
                audioPool[i] = Marshal.AllocHGlobal(BytesPerFrame);
            }

            try
            {
                this.IsSpeaking = true;
                int frameIndex = 0;
                while (offset < pcmData.Length && !token.IsCancellationRequested)
                {
                    if (this.IsMuted)
                    {
                        // Fill with silence if muted
                        Array.Clear(frameBuffer, 0, frameBuffer.Length);
                    }
                    else
                    {
                        int bytesToCopy = Math.Min(BytesPerFrame, pcmData.Length - offset);
                        Array.Copy(pcmData, offset, frameBuffer, 0, bytesToCopy);
                        if (bytesToCopy < BytesPerFrame)
                        {
                            Array.Clear(frameBuffer, bytesToCopy, BytesPerFrame - bytesToCopy);
                        }
                    }

                    if (this.IsAudioSendActive)
                    {
                        int slot = frameIndex % poolSize;
                        Marshal.Copy(frameBuffer, 0, audioPool[slot], BytesPerFrame);

                        var audioMediaBuffer = new SafeAudioMediaBuffer(
                            audioPool[slot],
                            (long)BytesPerFrame,
                            AudioFormat.Pcm16K,
                            (long)timestamp);

                        this.audioSocket.Send(audioMediaBuffer);
                    }

                    frameIndex++;
                    offset += BytesPerFrame;
                    timestamp += FrameDurationMs;
                    targetElapsedMs += FrameDurationMs;

                    // Drift-compensated pacing: ensure exactly 20ms of audio per 20ms real time
                    long currentElapsedMs = stopwatch.ElapsedMilliseconds;
                    int waitMs = (int)(targetElapsedMs - currentElapsedMs);
                    if (waitMs > 1)
                    {
                        if (waitMs > 15)
                        {
                            await Task.Delay(waitMs - 8, token).ConfigureAwait(false);
                        }
                        while (stopwatch.ElapsedMilliseconds < targetElapsedMs && !token.IsCancellationRequested)
                        {
                            Thread.SpinWait(100);
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                this.graphLogger?.Info("AudioSender: Audio playback was canceled.");
            }
            catch (Exception ex)
            {
                this.graphLogger?.Error(ex, "AudioSender: Error during audio frame playback.");
            }
            finally
            {
                this.IsSpeaking = false;
                // Brief delay so native audio transport finishes transmitting before memory is freed
                try { await Task.Delay(300).ConfigureAwait(false); } catch { }
                for (int i = 0; i < poolSize; i++)
                {
                    if (audioPool[i] != IntPtr.Zero)
                    {
                        Marshal.FreeHGlobal(audioPool[i]);
                        audioPool[i] = IntPtr.Zero;
                    }
                }
            }
        }

        /// <summary>
        /// Generates and plays a pleasant greeting chime (sine wave tone).
        /// </summary>
        public async Task PlayChimeAsync(CancellationToken cancellationToken = default)
        {
            int durationMs = 800;
            int totalSamples = (SampleRate * durationMs) / 1000;
            var pcm = new byte[totalSamples * 2];

            // Harmonic chord: D5 (587 Hz) then F#5 (740 Hz) then A5 (880 Hz)
            int chordDuration = totalSamples / 3;

            for (int i = 0; i < totalSamples; i++)
            {
                double t = (double)i / SampleRate;
                double freq = (i < chordDuration) ? 587.33 : (i < chordDuration * 2 ? 739.99 : 880.00);

                // Apply envelope (fade in and fade out)
                double env = Math.Sin(Math.PI * i / totalSamples);
                double sample = Math.Sin(2.0 * Math.PI * freq * t) * 0.4 * env;

                short sample16 = (short)(sample * short.MaxValue);
                pcm[i * 2] = (byte)(sample16 & 0xFF);
                pcm[i * 2 + 1] = (byte)((sample16 >> 8) & 0xFF);
            }

            await PlayPcmBytesAsync(pcm, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Synthesizes spoken voice from text using Windows System.Speech and plays it into the meeting.
        /// </summary>
        public async Task SpeakAsync(string text, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(text) || this.audioSocket == null)
            {
                return;
            }

            try
            {
                byte[] pcmData = null;
                using (var synth = new System.Speech.Synthesis.SpeechSynthesizer())
                using (var ms = new MemoryStream())
                {
                    synth.Rate = 1; // Natural, clear and energetic tempo
                    SelectConfiguredVoice(synth, this.graphLogger);
                    var format = new System.Speech.AudioFormat.SpeechAudioFormatInfo(SampleRate, System.Speech.AudioFormat.AudioBitsPerSample.Sixteen, System.Speech.AudioFormat.AudioChannel.Mono);
                    synth.SetOutputToAudioStream(ms, format);
                    synth.Speak(text);
                    pcmData = ms.ToArray();
                }

                if (pcmData != null && pcmData.Length > 0)
                {
                    this.graphLogger?.Info($"[AudioSender] Speaking text: {text} ({pcmData.Length} bytes)");
                    Console.WriteLine($">>> [AudioSender] Speaking into call: \"{text}\"");
                    await PlayPcmBytesAsync(pcmData, cancellationToken).ConfigureAwait(false);
                    Console.WriteLine($">>> [AudioSender] Finished speaking aloud into call: \"{text}\"");
                }
            }
            catch (Exception ex)
            {
                this.graphLogger?.Error(ex, $"[AudioSender] Failed to synthesize or play speech for '{text}'");
            }
        }

        /// <summary>
        /// Plays a welcome chime followed by a verbal introduction.
        /// </summary>
        public async Task PlayGreetingAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                await PlayChimeAsync(cancellationToken).ConfigureAwait(false);
                await Task.Delay(400, cancellationToken).ConfigureAwait(false);
                string greeting = TeamsCallingBot.Config.BotOptions.Current?.GreetingText;
                if (string.IsNullOrWhiteSpace(greeting))
                {
                    greeting = "Hello! Teams AI Assistant is now active and listening to this meeting.";
                }
                await SpeakAsync(greeting, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.graphLogger?.Error(ex, "[AudioSender] Error playing greeting.");
            }
        }

        private static bool voicesLogged;

        /// <summary>
        /// Picks the TTS voice: exact Bot:TtsVoiceName if installed, otherwise the first installed
        /// voice whose culture matches Bot:TtsCulture (default en-IN = Indian English), preferring Bot:TtsVoiceGender.
        /// </summary>
        internal static void SelectConfiguredVoice(System.Speech.Synthesis.SpeechSynthesizer synth, IGraphLogger logger)
        {
            var options = TeamsCallingBot.Config.BotOptions.Current;
            string wantedName = options?.TtsVoiceName;
            string wantedCulture = string.IsNullOrWhiteSpace(options?.TtsCulture) ? "en-IN" : options.TtsCulture;
            bool wantMale = string.Equals(options?.TtsVoiceGender, "Male", StringComparison.OrdinalIgnoreCase);

            try
            {
                var installed = synth.GetInstalledVoices();
                if (!voicesLogged)
                {
                    voicesLogged = true;
                    foreach (var v in installed)
                    {
                        var line = $"[TTS] Installed voice: {v.VoiceInfo.Name} ({v.VoiceInfo.Culture.Name}, {v.VoiceInfo.Gender}, enabled={v.Enabled})";
                        logger?.Info(line);
                        Console.WriteLine(">>> " + line);
                    }
                }

                if (!string.IsNullOrWhiteSpace(wantedName))
                {
                    foreach (var v in installed)
                    {
                        if (v.Enabled && string.Equals(v.VoiceInfo.Name, wantedName, StringComparison.OrdinalIgnoreCase))
                        {
                            synth.SelectVoice(v.VoiceInfo.Name);
                            return;
                        }
                    }

                    logger?.Warn($"[TTS] Configured voice '{wantedName}' is not installed - falling back to culture match.");
                }

                System.Speech.Synthesis.InstalledVoice best = null;
                foreach (var v in installed)
                {
                    if (!v.Enabled || !string.Equals(v.VoiceInfo.Culture.Name, wantedCulture, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    bool isMale = v.VoiceInfo.Gender == System.Speech.Synthesis.VoiceGender.Male;
                    if (best == null || isMale == wantMale)
                    {
                        best = v;
                        if (isMale == wantMale)
                        {
                            break;
                        }
                    }
                }

                if (best != null)
                {
                    synth.SelectVoice(best.VoiceInfo.Name);
                    return;
                }

                // Last resort: let SAPI pick by hints
                synth.SelectVoiceByHints(
                    wantMale ? System.Speech.Synthesis.VoiceGender.Male : System.Speech.Synthesis.VoiceGender.Female,
                    System.Speech.Synthesis.VoiceAge.Adult,
                    0,
                    new System.Globalization.CultureInfo(wantedCulture));
            }
            catch (Exception ex)
            {
                logger?.Warn($"[TTS] Voice selection failed ({ex.Message}) - using default voice.");
            }
        }

        public void Stop()
        {
            lock (this.playbackLock)
            {
                this.currentPlaybackCts?.Cancel();
            }
        }

        public void Dispose()
        {
            this.Stop();
            if (this.audioSocket != null)
            {
                this.audioSocket.AudioSendStatusChanged -= this.OnAudioSendStatusChanged;
            }
        }

        private static byte[] ExtractPcmFromWav(string wavFilePath)
        {
            using (var stream = File.OpenRead(wavFilePath))
            using (var reader = new BinaryReader(stream))
            {
                // Check RIFF header
                var riff = new string(reader.ReadChars(4));
                if (riff != "RIFF") return null;

                reader.ReadInt32(); // chunk size
                var wave = new string(reader.ReadChars(4));
                if (wave != "WAVE") return null;

                // Find data chunk
                while (stream.Position < stream.Length)
                {
                    var chunkId = new string(reader.ReadChars(4));
                    var chunkSize = reader.ReadInt32();

                    if (chunkId == "data")
                    {
                        return reader.ReadBytes(chunkSize);
                    }

                    stream.Seek(chunkSize, SeekOrigin.Current);
                }
            }

            return null;
        }
    }
}
