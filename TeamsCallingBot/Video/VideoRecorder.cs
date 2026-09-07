namespace TeamsCallingBot.Video
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Skype.Bots.Media;

    /// <summary>
    /// Records ONE incoming video source (a participant's screen share, or a participant's camera)
    /// to disk as MJPEG AVI segments, and periodically saves still JPEG snapshots for the minutes
    /// of meeting. One instance per (participant, stream) - the CallHandler keys them by the media
    /// source id (MSI) carried on every <see cref="VideoMediaBuffer"/>.
    ///
    /// Threading: <see cref="OnFrame"/> is called on the media platform's callback thread and MUST
    /// return quickly (the SDK drops/stalls if callbacks block). It only copies the raw bytes into a
    /// "latest frame" slot and signals the worker. The worker thread does colour conversion, JPEG
    /// encoding and disk writes. If frames arrive faster than the worker can process them, older
    /// pending frames are simply replaced (we sample to the target fps anyway).
    /// </summary>
    public sealed class VideoRecorder : IDisposable
    {
        private readonly object frameLock = new object();
        private readonly AutoResetEvent frameSignal = new AutoResetEvent(false);
        private readonly Action<string> log;
        private readonly List<string> segmentPaths = new List<string>();
        private readonly List<SnapshotInfo> snapshots = new List<SnapshotInfo>();
        private readonly Stopwatch clock = Stopwatch.StartNew();

        private PendingFrame pending;
        private Thread worker;
        private volatile bool stopping;
        private MjpegAviWriter aviWriter;
        private int segmentNumber;
        private double segmentStartSeconds;
        private double lastSnapshotSeconds = double.NegativeInfinity;
        private Bitmap latestBitmap;
        private byte[] latestJpeg;
        private long lastFlushTicks;
        private int disposed;

        public VideoRecorder(string outputDirectory, string fileStem, uint mediaSourceId, string sourceLabel, VideoRecorderSettings settings, Action<string> log)
        {
            this.OutputDirectory = outputDirectory ?? throw new ArgumentNullException(nameof(outputDirectory));
            this.FileStem = SanitizeFileName(fileStem ?? "video");
            this.MediaSourceId = mediaSourceId;
            this.SourceLabel = sourceLabel ?? $"MSI {mediaSourceId}";
            this.Settings = settings ?? new VideoRecorderSettings();
            this.log = log ?? (_ => { });
            this.StartedAt = DateTime.Now;
        }

        public string OutputDirectory { get; }

        public string FileStem { get; }

        public uint MediaSourceId { get; }

        public string SourceLabel { get; set; }

        public VideoRecorderSettings Settings { get; }

        public DateTime StartedAt { get; }

        public DateTime? StoppedAt { get; private set; }

        public int Width { get; private set; }

        public int Height { get; private set; }

        public long FramesReceived { get; private set; }

        public long FramesWritten { get; private set; }

        public long FramesDropped { get; private set; }

        public bool IsRunning => this.worker != null && !this.stopping;

        public IReadOnlyList<string> SegmentPaths
        {
            get
            {
                lock (this.segmentPaths)
                {
                    return this.segmentPaths.ToArray();
                }
            }
        }

        public IReadOnlyList<SnapshotInfo> Snapshots
        {
            get
            {
                lock (this.snapshots)
                {
                    return this.snapshots.ToArray();
                }
            }
        }

        /// <summary>Raised on the worker thread whenever a snapshot JPEG has been saved.</summary>
        public event Action<VideoRecorder, SnapshotInfo> SnapshotSaved;

        public void Start()
        {
            if (this.worker != null)
            {
                return;
            }

            this.worker = new Thread(this.WorkerLoop)
            {
                IsBackground = true,
                Name = $"VideoRecorder-{this.FileStem}",
                Priority = ThreadPriority.BelowNormal,
            };
            this.worker.Start();
            this.log($"[VideoRecorder] Started for {this.SourceLabel} (MSI {this.MediaSourceId}) -> {Path.Combine(this.OutputDirectory, this.FileStem)}_*.avi @ {this.Settings.Fps} fps");
        }

        /// <summary>
        /// Called on the media callback thread. Copies the frame bytes and returns immediately.
        /// The caller still owns (and must dispose) the buffer.
        /// </summary>
        public void OnFrame(VideoMediaBuffer buffer)
        {
            if (buffer == null || buffer.Data == IntPtr.Zero || this.stopping || this.worker == null)
            {
                return;
            }

            var format = buffer.VideoFormat ?? buffer.OriginalVideoFormat;
            if (format == null || format.Width <= 0 || format.Height <= 0)
            {
                return;
            }

            int width = format.Width;
            int height = format.Height;
            int stride = buffer.Stride > 0 ? buffer.Stride : width;
            long expected;
            if (format.VideoColorFormat == VideoColorFormat.NV12)
            {
                expected = (long)Math.Max(stride, width) * height * 3 / 2;
            }
            else if (format.VideoColorFormat == VideoColorFormat.Rgb24)
            {
                expected = (long)Math.Max(stride, width * 3) * height;
            }
            else
            {
                return; // H264 / Yuy2 are never requested by this bot
            }

            long available = buffer.Length > 0 ? Math.Min(buffer.Length, expected) : expected;
            if (available <= 0 || available > int.MaxValue)
            {
                return;
            }

            this.FramesReceived++;

            PendingFrame frame;
            lock (this.frameLock)
            {
                // Reuse the previous pending frame's array when it is big enough and was never consumed.
                frame = this.pending;
                if (frame != null)
                {
                    this.FramesDropped++;
                }

                if (frame == null || frame.Data.Length < available)
                {
                    frame = new PendingFrame { Data = new byte[available] };
                }

                frame.Length = (int)available;
                frame.Width = width;
                frame.Height = height;
                frame.Stride = stride;
                frame.ColorFormat = format.VideoColorFormat;
                frame.ElapsedSeconds = this.clock.Elapsed.TotalSeconds;
                frame.ArrivedAt = DateTime.Now;
                Marshal.Copy(buffer.Data, frame.Data, 0, frame.Length);
                this.pending = frame;
            }

            this.frameSignal.Set();
        }

        /// <summary>Returns a copy of the most recent decoded frame, or null.</summary>
        public Bitmap GetLatestBitmapCopy()
        {
            lock (this.frameLock)
            {
                return this.latestBitmap == null ? null : (Bitmap)this.latestBitmap.Clone();
            }
        }

        /// <summary>Saves the most recent frame as a JPEG to the given path (on demand snapshot).</summary>
        public string SaveLatestFrame(string path)
        {
            byte[] jpeg;
            lock (this.frameLock)
            {
                jpeg = this.latestJpeg;
            }

            if (jpeg == null)
            {
                return null;
            }

            File.WriteAllBytes(path, jpeg);
            return path;
        }

        /// <summary>Stops recording, finalises the current AVI segment and returns once the worker exited.</summary>
        public void Stop()
        {
            if (this.worker == null || this.stopping)
            {
                return;
            }

            this.stopping = true;
            this.frameSignal.Set();
            if (!this.worker.Join(TimeSpan.FromSeconds(15)))
            {
                this.log($"[VideoRecorder] Worker for {this.SourceLabel} did not exit in time.");
            }

            this.StoppedAt = DateTime.Now;
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref this.disposed, 1) == 1)
            {
                return;
            }

            this.Stop();
            lock (this.frameLock)
            {
                this.latestBitmap?.Dispose();
                this.latestBitmap = null;
            }

            this.frameSignal.Dispose();
        }

        // ------------------------------------------------------------------

        private void WorkerLoop()
        {
            try
            {
                while (true)
                {
                    this.frameSignal.WaitOne(500);

                    PendingFrame frame;
                    lock (this.frameLock)
                    {
                        frame = this.pending;
                        this.pending = null;
                    }

                    if (frame != null)
                    {
                        try
                        {
                            this.ProcessFrame(frame);
                        }
                        catch (Exception ex)
                        {
                            this.log($"[VideoRecorder] Frame processing error for {this.SourceLabel}: {ex.Message}");
                        }
                    }
                    else if (this.aviWriter != null && !this.stopping)
                    {
                        // Keep the timeline moving during static periods so playback duration matches
                        // wall-clock even when Teams sends no frames for a long time.
                        int idx = this.FrameIndexFor(this.clock.Elapsed.TotalSeconds) - 1;
                        if (idx > this.aviWriter.FrameCount)
                        {
                            this.aviWriter.PadTo(idx);
                        }

                        this.MaybeFlush();
                    }

                    if (this.stopping)
                    {
                        // Drain the very last frame (if one arrived between the flag and the signal).
                        lock (this.frameLock)
                        {
                            frame = this.pending;
                            this.pending = null;
                        }

                        if (frame != null)
                        {
                            try { this.ProcessFrame(frame); } catch { }
                        }

                        break;
                    }
                }
            }
            finally
            {
                this.CloseSegment(final: true);
            }
        }

        private void ProcessFrame(PendingFrame frame)
        {
            Bitmap bitmap = frame.ColorFormat == VideoColorFormat.NV12
                ? VideoFrameConverter.ConvertNV12ToBitmap(frame.Data, frame.Width, frame.Height, frame.Stride)
                : VideoFrameConverter.ConvertRGB24ToBitmap(frame.Data, frame.Width, frame.Height, frame.Stride);

            if (bitmap == null)
            {
                return;
            }

            try
            {
                if (this.aviWriter == null)
                {
                    this.OpenSegment(bitmap.Width, bitmap.Height, frame.ElapsedSeconds);
                }
                else if (bitmap.Width != this.Width || bitmap.Height != this.Height)
                {
                    // Teams changed the share resolution mid-stream; AVI dimensions are fixed so scale.
                    var resized = VideoFrameConverter.ResizeBitmap(bitmap, this.Width, this.Height);
                    bitmap.Dispose();
                    bitmap = resized;
                }

                byte[] jpeg = VideoFrameConverter.EncodeJpeg(bitmap, this.Settings.JpegQuality);
                int frameIndex = this.FrameIndexFor(frame.ElapsedSeconds);

                if (this.aviWriter.WriteFrame(jpeg, frameIndex))
                {
                    this.FramesWritten++;
                }

                lock (this.frameLock)
                {
                    this.latestBitmap?.Dispose();
                    this.latestBitmap = bitmap;
                    this.latestJpeg = jpeg;
                    bitmap = null; // ownership transferred
                }

                if (frame.ElapsedSeconds - this.lastSnapshotSeconds >= this.Settings.SnapshotIntervalSeconds)
                {
                    this.lastSnapshotSeconds = frame.ElapsedSeconds;
                    this.SaveSnapshot(jpeg, frame.ArrivedAt);
                }

                this.MaybeFlush();

                if (this.aviWriter.FileLength >= this.Settings.MaxSegmentBytes)
                {
                    this.CloseSegment(final: false);
                    this.OpenSegment(this.Width, this.Height, this.clock.Elapsed.TotalSeconds);
                }
            }
            finally
            {
                bitmap?.Dispose();
            }
        }

        private int FrameIndexFor(double elapsedSeconds)
        {
            double relative = Math.Max(0, elapsedSeconds - this.segmentStartSeconds);
            return (int)Math.Round(relative * this.Settings.Fps);
        }

        private void OpenSegment(int width, int height, double startSeconds)
        {
            this.segmentNumber++;
            this.Width = width;
            this.Height = height;
            this.segmentStartSeconds = startSeconds;
            string path = Path.Combine(this.OutputDirectory, $"{this.FileStem}_part{this.segmentNumber:D2}.avi");
            this.aviWriter = new MjpegAviWriter(path, width, height, this.Settings.Fps);
            lock (this.segmentPaths)
            {
                this.segmentPaths.Add(path);
            }

            this.lastFlushTicks = Environment.TickCount;
            this.log($"[VideoRecorder] Segment opened: {path} ({width}x{height} @ {this.Settings.Fps} fps)");
        }

        private void CloseSegment(bool final)
        {
            var writer = this.aviWriter;
            if (writer == null)
            {
                return;
            }

            this.aviWriter = null;
            try
            {
                if (final)
                {
                    int idx = this.FrameIndexFor(this.clock.Elapsed.TotalSeconds) - 1;
                    if (idx > writer.FrameCount)
                    {
                        writer.PadTo(idx);
                    }
                }

                writer.Close();
                this.log($"[VideoRecorder] Segment closed: {writer.Path} ({writer.RealFrameCount} frames, {writer.FrameCount} slots, {writer.FileLength / 1024 / 1024} MB)");
            }
            catch (Exception ex)
            {
                this.log($"[VideoRecorder] Error closing segment {writer.Path}: {ex.Message}");
            }
        }

        private void MaybeFlush()
        {
            if (this.aviWriter == null)
            {
                return;
            }

            if (Environment.TickCount - this.lastFlushTicks > 5000)
            {
                this.lastFlushTicks = Environment.TickCount;
                try
                {
                    this.aviWriter.Flush();
                }
                catch (Exception ex)
                {
                    this.log($"[VideoRecorder] Flush error: {ex.Message}");
                }
            }
        }

        private void SaveSnapshot(byte[] jpeg, DateTime at)
        {
            try
            {
                string path = Path.Combine(this.OutputDirectory, $"{this.Settings.SnapshotPrefix}_{this.FileStem}_{at:yyyyMMdd_HHmmss_fff}.jpg");
                File.WriteAllBytes(path, jpeg);
                var info = new SnapshotInfo { Path = path, CapturedAt = at, SourceLabel = this.SourceLabel, MediaSourceId = this.MediaSourceId };
                lock (this.snapshots)
                {
                    this.snapshots.Add(info);
                }

                this.SnapshotSaved?.Invoke(this, info);
            }
            catch (Exception ex)
            {
                this.log($"[VideoRecorder] Snapshot error: {ex.Message}");
            }
        }

        /// <summary>
        /// Optionally converts finished AVI segments to MP4 with an external ffmpeg. Returns the list
        /// of MP4 paths produced. Safe to call when ffmpeg is not configured (returns empty list).
        /// </summary>
        public static async Task<List<string>> ConvertSegmentsToMp4Async(IEnumerable<string> aviPaths, string ffmpegPath, Action<string> log)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(ffmpegPath) || !File.Exists(ffmpegPath))
            {
                return result;
            }

            foreach (var avi in aviPaths)
            {
                if (!File.Exists(avi))
                {
                    continue;
                }

                string mp4 = Path.ChangeExtension(avi, ".mp4");
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = ffmpegPath,
                        Arguments = $"-y -loglevel error -i \"{avi}\" -c:v libx264 -pix_fmt yuv420p -preset veryfast -crf 23 -movflags +faststart \"{mp4}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardError = true,
                    };

                    using (var process = new Process { StartInfo = psi, EnableRaisingEvents = true })
                    {
                        var tcs = new TaskCompletionSource<bool>();
                        process.Exited += (s, e) => tcs.TrySetResult(true);
                        process.Start();
                        string stderr = await process.StandardError.ReadToEndAsync().ConfigureAwait(false);
                        await tcs.Task.ConfigureAwait(false);
                        if (process.ExitCode == 0 && File.Exists(mp4))
                        {
                            result.Add(mp4);
                            log?.Invoke($"[VideoRecorder] Converted {Path.GetFileName(avi)} -> {Path.GetFileName(mp4)}");
                        }
                        else
                        {
                            log?.Invoke($"[VideoRecorder] ffmpeg failed for {avi} (exit {process.ExitCode}): {stderr}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.Invoke($"[VideoRecorder] ffmpeg error for {avi}: {ex.Message}");
                }
            }

            return result;
        }

        public static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "unknown";
            }

            var invalid = Path.GetInvalidFileNameChars();
            var chars = name.Trim().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (Array.IndexOf(invalid, chars[i]) >= 0 || chars[i] == ' ' || chars[i] == '.')
                {
                    chars[i] = '_';
                }
            }

            var cleaned = new string(chars);
            return cleaned.Length > 60 ? cleaned.Substring(0, 60) : cleaned;
        }

        private sealed class PendingFrame
        {
            public byte[] Data;
            public int Length;
            public int Width;
            public int Height;
            public int Stride;
            public VideoColorFormat ColorFormat;
            public double ElapsedSeconds;
            public DateTime ArrivedAt;
        }
    }

    public sealed class VideoRecorderSettings
    {
        public int Fps { get; set; } = 5;

        public int JpegQuality { get; set; } = 75;

        public int SnapshotIntervalSeconds { get; set; } = 10;

        public long MaxSegmentBytes { get; set; } = 1_500_000_000;

        public string SnapshotPrefix { get; set; } = "03_snapshot";
    }

    public sealed class SnapshotInfo
    {
        public string Path { get; set; }

        public DateTime CapturedAt { get; set; }

        public string SourceLabel { get; set; }

        public uint MediaSourceId { get; set; }
    }
}
