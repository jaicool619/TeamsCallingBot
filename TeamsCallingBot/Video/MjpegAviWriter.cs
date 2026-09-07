namespace TeamsCallingBot.Video
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Minimal, dependency-free writer for Motion-JPEG AVI files (RIFF 'AVI ' with a single 'vids'
    /// stream, fourcc MJPG, plus an idx1 index). Every mainstream player (VLC, Windows Media Player,
    /// Films &amp; TV, ffmpeg, browsers via conversion) opens these files.
    ///
    /// Why MJPEG/AVI and not MP4/H.264: encoding H.264 in-process on net472 needs either Media
    /// Foundation interop or a native library - both fragile on a media-bot VM. MJPEG only needs the
    /// GDI+ JPEG encoder that this project already uses for photos. If FfmpegPath is configured the
    /// finished AVI is additionally converted to MP4 at call end (see VideoRecorder).
    ///
    /// Variable-rate input, fixed-rate output: Teams sends screen-share frames only when the content
    /// changes. Callers therefore pass an absolute frame index (derived from wall-clock time); gaps
    /// are filled with zero-length '00dc' chunks, which the AVI spec defines as "repeat the previous
    /// frame". That keeps playback timing real-time without storing duplicate JPEGs.
    ///
    /// Crash resilience: the RIFF/LIST sizes and frame counts in the header are patched on every
    /// <see cref="Flush"/> so a file left behind by a killed process is still mostly playable
    /// (players rebuild the index by scanning the movi list).
    /// </summary>
    public sealed class MjpegAviWriter : IDisposable
    {
        private const uint AVIF_HASINDEX = 0x00000010;
        private const uint AVIF_TRUSTCKTYPE = 0x00000800;
        private const uint AVIIF_KEYFRAME = 0x00000010;

        private readonly FileStream stream;
        private readonly BinaryWriter writer;
        private readonly List<IndexEntry> index = new List<IndexEntry>();
        private readonly object sync = new object();

        private long riffSizePos;
        private long avihTotalFramesPos;
        private long strhLengthPos;
        private long moviListSizePos;
        private long moviDataStart;
        private int lastFrameIndex = -1;
        private uint maxChunkSize;
        private bool closed;

        public MjpegAviWriter(string path, int width, int height, int fps)
        {
            if (width <= 0 || height <= 0)
            {
                throw new ArgumentException("Width and height must be positive.");
            }

            this.Path = path;
            this.Width = width;
            this.Height = height;
            this.Fps = Math.Max(1, Math.Min(60, fps));

            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path) ?? ".");
            this.stream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.Read, 1 << 16);
            this.writer = new BinaryWriter(this.stream);
            this.WriteHeaders();
        }

        public string Path { get; }

        public int Width { get; }

        public int Height { get; }

        public int Fps { get; }

        /// <summary>Total frame slots written so far (real + repeated).</summary>
        public int FrameCount => this.lastFrameIndex + 1;

        /// <summary>Number of real JPEG frames written (excludes repeat markers).</summary>
        public int RealFrameCount { get; private set; }

        public long FileLength => this.stream.Length;

        public bool IsClosed => this.closed;

        /// <summary>
        /// Writes a JPEG frame at the given absolute frame index. Frame slots between the previous
        /// frame and this one are filled with "repeat previous frame" markers. If the index is not
        /// beyond the last written frame the frame is dropped (the caller is sampling faster than
        /// the output frame rate) and false is returned.
        /// </summary>
        public bool WriteFrame(byte[] jpeg, int frameIndex)
        {
            if (jpeg == null || jpeg.Length == 0)
            {
                return false;
            }

            lock (this.sync)
            {
                if (this.closed)
                {
                    return false;
                }

                if (frameIndex <= this.lastFrameIndex)
                {
                    return false;
                }

                // Fill the gap with empty (repeat) chunks.
                for (int i = this.lastFrameIndex + 1; i < frameIndex; i++)
                {
                    this.WriteChunk(null);
                }

                this.WriteChunk(jpeg);
                this.lastFrameIndex = frameIndex;
                this.RealFrameCount++;
                return true;
            }
        }

        /// <summary>Extends the timeline with repeat markers up to (and including) the given frame index.</summary>
        public void PadTo(int frameIndex)
        {
            lock (this.sync)
            {
                if (this.closed)
                {
                    return;
                }

                while (this.lastFrameIndex < frameIndex)
                {
                    this.WriteChunk(null);
                    this.lastFrameIndex++;
                }
            }
        }

        /// <summary>Patches header sizes/counts and flushes to disk without closing.</summary>
        public void Flush()
        {
            lock (this.sync)
            {
                if (this.closed)
                {
                    return;
                }

                this.PatchSizes(includeIndex: false);
                this.stream.Flush(true);
            }
        }

        /// <summary>Writes the idx1 index, finalises all sizes and closes the file.</summary>
        public void Close()
        {
            lock (this.sync)
            {
                if (this.closed)
                {
                    return;
                }

                this.closed = true;

                this.stream.Seek(0, SeekOrigin.End);
                long moviEnd = this.stream.Position;

                // idx1
                this.WriteFourCC("idx1");
                this.writer.Write((uint)(this.index.Count * 16));
                foreach (var entry in this.index)
                {
                    this.WriteFourCC("00dc");
                    this.writer.Write(entry.Flags);
                    this.writer.Write(entry.Offset);
                    this.writer.Write(entry.Size);
                }

                this.PatchSizes(includeIndex: true, moviEnd: moviEnd);
                this.writer.Flush();
                this.stream.Flush(true);
                this.writer.Dispose();
                this.stream.Dispose();
            }
        }

        public void Dispose()
        {
            this.Close();
        }

        // ------------------------------------------------------------------

        private void WriteChunk(byte[] jpeg)
        {
            this.stream.Seek(0, SeekOrigin.End);
            uint size = jpeg == null ? 0u : (uint)jpeg.Length;

            // Offset in idx1 is relative to the 'movi' fourcc position (position of the fourcc itself
            // is moviDataStart - 4). Most players accept either convention; this one matches ffmpeg/VLC.
            uint offset = (uint)(this.stream.Position - (this.moviDataStart - 4));

            this.WriteFourCC("00dc");
            this.writer.Write(size);
            if (jpeg != null)
            {
                this.writer.Write(jpeg);
                if ((size & 1) == 1)
                {
                    this.writer.Write((byte)0); // RIFF chunks are word aligned
                }

                if (size > this.maxChunkSize)
                {
                    this.maxChunkSize = size;
                }
            }

            this.index.Add(new IndexEntry
            {
                Flags = jpeg == null ? 0u : AVIIF_KEYFRAME,
                Offset = offset,
                Size = size,
            });
        }

        private void WriteHeaders()
        {
            uint microSecPerFrame = (uint)(1_000_000.0 / this.Fps);

            this.WriteFourCC("RIFF");
            this.riffSizePos = this.stream.Position;
            this.writer.Write(0u); // patched later
            this.WriteFourCC("AVI ");

            // LIST hdrl
            this.WriteFourCC("LIST");
            long hdrlSizePos = this.stream.Position;
            this.writer.Write(0u);
            long hdrlStart = this.stream.Position;
            this.WriteFourCC("hdrl");

            // avih
            this.WriteFourCC("avih");
            this.writer.Write(56u);
            this.writer.Write(microSecPerFrame);                       // dwMicroSecPerFrame
            this.writer.Write((uint)(this.Width * this.Height * 3 * this.Fps)); // dwMaxBytesPerSec (upper bound)
            this.writer.Write(0u);                                     // dwPaddingGranularity
            this.writer.Write(AVIF_HASINDEX | AVIF_TRUSTCKTYPE);       // dwFlags
            this.avihTotalFramesPos = this.stream.Position;
            this.writer.Write(0u);                                     // dwTotalFrames (patched)
            this.writer.Write(0u);                                     // dwInitialFrames
            this.writer.Write(1u);                                     // dwStreams
            this.writer.Write((uint)(this.Width * this.Height * 3));   // dwSuggestedBufferSize
            this.writer.Write((uint)this.Width);
            this.writer.Write((uint)this.Height);
            this.writer.Write(0u); this.writer.Write(0u); this.writer.Write(0u); this.writer.Write(0u); // dwReserved[4]

            // LIST strl
            this.WriteFourCC("LIST");
            long strlSizePos = this.stream.Position;
            this.writer.Write(0u);
            long strlStart = this.stream.Position;
            this.WriteFourCC("strl");

            // strh
            this.WriteFourCC("strh");
            this.writer.Write(56u);
            this.WriteFourCC("vids");                                  // fccType
            this.WriteFourCC("MJPG");                                  // fccHandler
            this.writer.Write(0u);                                     // dwFlags
            this.writer.Write((ushort)0);                              // wPriority
            this.writer.Write((ushort)0);                              // wLanguage
            this.writer.Write(0u);                                     // dwInitialFrames
            this.writer.Write(1u);                                     // dwScale
            this.writer.Write((uint)this.Fps);                         // dwRate
            this.writer.Write(0u);                                     // dwStart
            this.strhLengthPos = this.stream.Position;
            this.writer.Write(0u);                                     // dwLength (patched)
            this.writer.Write((uint)(this.Width * this.Height * 3));   // dwSuggestedBufferSize
            this.writer.Write(0xFFFFFFFFu);                            // dwQuality (default)
            this.writer.Write(0u);                                     // dwSampleSize
            this.writer.Write((short)0); this.writer.Write((short)0);  // rcFrame left, top
            this.writer.Write((short)this.Width); this.writer.Write((short)this.Height); // right, bottom

            // strf (BITMAPINFOHEADER)
            this.WriteFourCC("strf");
            this.writer.Write(40u);
            this.writer.Write(40u);                                    // biSize
            this.writer.Write(this.Width);                             // biWidth
            this.writer.Write(this.Height);                            // biHeight
            this.writer.Write((ushort)1);                              // biPlanes
            this.writer.Write((ushort)24);                             // biBitCount
            this.WriteFourCC("MJPG");                                  // biCompression
            this.writer.Write((uint)(this.Width * this.Height * 3));   // biSizeImage
            this.writer.Write(0); this.writer.Write(0);                // biXPelsPerMeter, biYPelsPerMeter
            this.writer.Write(0u); this.writer.Write(0u);              // biClrUsed, biClrImportant

            long strlEnd = this.stream.Position;
            this.PatchUInt(strlSizePos, (uint)(strlEnd - strlStart));

            long hdrlEnd = this.stream.Position;
            this.PatchUInt(hdrlSizePos, (uint)(hdrlEnd - hdrlStart));

            // JUNK padding so the movi list starts at a 2 KB boundary (conventional, helps some players)
            long pad = 2048 - (this.stream.Position + 8) % 2048;
            if (pad < 8)
            {
                pad += 2048;
            }

            this.WriteFourCC("JUNK");
            this.writer.Write((uint)(pad - 8));
            this.writer.Write(new byte[pad - 8]);

            // LIST movi
            this.WriteFourCC("LIST");
            this.moviListSizePos = this.stream.Position;
            this.writer.Write(0u); // patched
            this.WriteFourCC("movi");
            this.moviDataStart = this.stream.Position;

            this.writer.Flush();
        }

        private void PatchSizes(bool includeIndex, long moviEnd = -1)
        {
            long end = this.stream.Length;
            if (moviEnd < 0)
            {
                moviEnd = end;
            }

            uint totalFrames = (uint)Math.Max(0, this.lastFrameIndex + 1);
            this.PatchUInt(this.riffSizePos, (uint)(end - 8));
            this.PatchUInt(this.avihTotalFramesPos, totalFrames);
            this.PatchUInt(this.strhLengthPos, totalFrames);
            this.PatchUInt(this.moviListSizePos, (uint)(moviEnd - (this.moviDataStart - 4)));
            this.stream.Seek(0, SeekOrigin.End);
        }

        private void PatchUInt(long position, uint value)
        {
            long current = this.stream.Position;
            this.stream.Seek(position, SeekOrigin.Begin);
            this.writer.Write(value);
            this.writer.Flush();
            this.stream.Seek(current, SeekOrigin.Begin);
        }

        private void WriteFourCC(string fourCC)
        {
            this.writer.Write(Encoding.ASCII.GetBytes(fourCC));
        }

        private struct IndexEntry
        {
            public uint Flags;
            public uint Offset;
            public uint Size;
        }
    }
}
