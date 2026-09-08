namespace TeamsCallingBot.Video
{
    using System;
    using System.Drawing;
    using System.Drawing.Drawing2D;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Video frame colour converter between Teams NV12 / RGB24 buffers and System.Drawing.Bitmap,
    /// plus the renderer for the bot's own outgoing "status card" video.
    ///
    /// Performance notes (2026-09-04): the previous implementation used per-pixel floating point
    /// math and a managed intermediate array, costing ~150 ms for a 1080p frame - far too slow to
    /// record screen share as video. The conversions below use fixed-point integer BT.601 math
    /// over raw pointers (AllowUnsafeBlocks is already on in the csproj) and run ~10x faster.
    /// The bot card is rendered once as a static background and only the dynamic parts (status
    /// pill, waveform, timestamp) are painted per frame.
    /// </summary>
    public static class VideoFrameConverter
    {
        private static readonly object CardLock = new object();
        private static Bitmap cachedCardBackground;
        private static string cachedCardBotName;
        private static ImageCodecInfo jpegCodec;

        // ------------------------------------------------------------------
        // NV12 / RGB24 -> Bitmap
        // ------------------------------------------------------------------

        /// <summary>
        /// Converts an NV12 buffer (full-res Y plane followed by interleaved, 2x2 subsampled UV plane)
        /// into a 24bpp BGR Bitmap. <paramref name="stride"/> is the Y-plane stride reported by the
        /// media platform; the UV plane is assumed to use the same stride (this is how the Teams
        /// media SDK lays out its receive buffers).
        /// </summary>
        public static Bitmap ConvertNV12ToBitmap(IntPtr data, int width, int height, int stride)
        {
            if (data == IntPtr.Zero || width <= 0 || height <= 0)
            {
                return null;
            }

            if (stride < width)
            {
                stride = width;
            }

            var bmp = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            var bmpData = bmp.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
            try
            {
                unsafe
                {
                    byte* src = (byte*)data.ToPointer();
                    byte* uvPlane = src + ((long)stride * height);
                    byte* dst = (byte*)bmpData.Scan0.ToPointer();

                    for (int y = 0; y < height; y++)
                    {
                        byte* yRow = src + ((long)y * stride);
                        byte* uvRow = uvPlane + ((long)(y >> 1) * stride);
                        byte* dstRow = dst + ((long)y * bmpData.Stride);

                        for (int x = 0; x < width; x++)
                        {
                            int c = yRow[x] - 16;
                            int uvIdx = (x >> 1) << 1;
                            int d = uvRow[uvIdx] - 128;
                            int e = uvRow[uvIdx + 1] - 128;

                            int c298 = 298 * c + 128;
                            int r = (c298 + 409 * e) >> 8;
                            int g = (c298 - 100 * d - 208 * e) >> 8;
                            int b = (c298 + 516 * d) >> 8;

                            byte* px = dstRow + (x * 3);
                            px[0] = (byte)(b < 0 ? 0 : (b > 255 ? 255 : b));
                            px[1] = (byte)(g < 0 ? 0 : (g > 255 ? 255 : g));
                            px[2] = (byte)(r < 0 ? 0 : (r > 255 ? 255 : r));
                        }
                    }
                }
            }
            finally
            {
                bmp.UnlockBits(bmpData);
            }

            return bmp;
        }

        /// <summary>Managed-array overload used by the recorder after the frame was copied off the media thread.</summary>
        public static Bitmap ConvertNV12ToBitmap(byte[] nv12, int width, int height, int stride)
        {
            if (nv12 == null || nv12.Length == 0)
            {
                return null;
            }

            long required = (long)Math.Max(stride, width) * height * 3 / 2;
            if (nv12.Length < required)
            {
                return null;
            }

            var handle = GCHandle.Alloc(nv12, GCHandleType.Pinned);
            try
            {
                return ConvertNV12ToBitmap(handle.AddrOfPinnedObject(), width, height, stride);
            }
            finally
            {
                handle.Free();
            }
        }

        /// <summary>Converts an RGB24 (actually BGR byte order, as GDI+ expects) buffer into a 24bpp Bitmap.</summary>
        public static Bitmap ConvertRGB24ToBitmap(IntPtr data, int width, int height, int stride)
        {
            if (data == IntPtr.Zero || width <= 0 || height <= 0)
            {
                return null;
            }

            if (stride < width * 3)
            {
                stride = width * 3;
            }

            var bmp = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            var bmpData = bmp.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
            try
            {
                unsafe
                {
                    int rowBytes = width * 3;
                    byte* src = (byte*)data.ToPointer();
                    byte* dst = (byte*)bmpData.Scan0.ToPointer();
                    for (int y = 0; y < height; y++)
                    {
                        Buffer.MemoryCopy(src + ((long)y * stride), dst + ((long)y * bmpData.Stride), rowBytes, rowBytes);
                    }
                }
            }
            finally
            {
                bmp.UnlockBits(bmpData);
            }

            return bmp;
        }

        public static Bitmap ConvertRGB24ToBitmap(byte[] rgb, int width, int height, int stride)
        {
            if (rgb == null || rgb.Length < (long)Math.Max(stride, width * 3) * height)
            {
                return null;
            }

            var handle = GCHandle.Alloc(rgb, GCHandleType.Pinned);
            try
            {
                return ConvertRGB24ToBitmap(handle.AddrOfPinnedObject(), width, height, stride);
            }
            finally
            {
                handle.Free();
            }
        }

        // ------------------------------------------------------------------
        // Bitmap -> NV12 (for the bot's outgoing video)
        // ------------------------------------------------------------------

        /// <summary>
        /// Converts a Bitmap to a tightly packed NV12 byte array of exactly width*height*3/2 bytes,
        /// resizing first if the bitmap dimensions differ.
        /// </summary>
        public static byte[] ConvertBitmapToNV12(Bitmap bitmap, int width, int height)
        {
            var nv12 = new byte[width * height * 3 / 2];
            Bitmap source = bitmap;
            bool disposeSource = false;

            if (bitmap.Width != width || bitmap.Height != height || bitmap.PixelFormat != PixelFormat.Format24bppRgb)
            {
                source = ResizeBitmap(bitmap, width, height);
                disposeSource = true;
            }

            try
            {
                var bmpData = source.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                try
                {
                    unsafe
                    {
                        byte* src = (byte*)bmpData.Scan0.ToPointer();
                        fixed (byte* dst = nv12)
                        {
                            byte* yPlane = dst;
                            byte* uvPlane = dst + (width * height);

                            for (int y = 0; y < height; y++)
                            {
                                byte* row = src + ((long)y * bmpData.Stride);
                                byte* yRow = yPlane + (y * width);
                                byte* uvRow = uvPlane + ((y >> 1) * width);
                                bool sampleUv = (y & 1) == 0;

                                for (int x = 0; x < width; x++)
                                {
                                    byte* px = row + (x * 3);
                                    int b = px[0];
                                    int g = px[1];
                                    int r = px[2];

                                    yRow[x] = (byte)(((66 * r + 129 * g + 25 * b + 128) >> 8) + 16);

                                    if (sampleUv && (x & 1) == 0)
                                    {
                                        int u = ((-38 * r - 74 * g + 112 * b + 128) >> 8) + 128;
                                        int v = ((112 * r - 94 * g - 18 * b + 128) >> 8) + 128;
                                        uvRow[x] = (byte)(u < 0 ? 0 : (u > 255 ? 255 : u));
                                        uvRow[x + 1] = (byte)(v < 0 ? 0 : (v > 255 ? 255 : v));
                                    }
                                }
                            }
                        }
                    }
                }
                finally
                {
                    source.UnlockBits(bmpData);
                }
            }
            finally
            {
                if (disposeSource)
                {
                    source.Dispose();
                }
            }

            return nv12;
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        public static Bitmap ResizeBitmap(Bitmap source, int width, int height)
        {
            var result = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(result))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBilinear;
                g.CompositingQuality = CompositingQuality.HighSpeed;
                g.SmoothingMode = SmoothingMode.HighSpeed;
                g.DrawImage(source, new Rectangle(0, 0, width, height));
            }

            return result;
        }

        /// <summary>Encodes a bitmap as JPEG bytes at the given quality (1-100).</summary>
        public static byte[] EncodeJpeg(Bitmap bitmap, int quality)
        {
            if (bitmap == null)
            {
                return null;
            }

            quality = Math.Max(1, Math.Min(100, quality));
            var codec = GetJpegCodec();
            using (var ms = new MemoryStream())
            {
                if (codec != null)
                {
                    using (var encParams = new EncoderParameters(1))
                    {
                        encParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)quality);
                        bitmap.Save(ms, codec, encParams);
                    }
                }
                else
                {
                    bitmap.Save(ms, ImageFormat.Jpeg);
                }

                return ms.ToArray();
            }
        }

        private static ImageCodecInfo GetJpegCodec()
        {
            if (jpegCodec != null)
            {
                return jpegCodec;
            }

            foreach (var codec in ImageCodecInfo.GetImageEncoders())
            {
                if (codec.FormatID == ImageFormat.Jpeg.Guid)
                {
                    jpegCodec = codec;
                    break;
                }
            }

            return jpegCodec;
        }

        // ------------------------------------------------------------------
        // Bot status card (outgoing video)
        // ------------------------------------------------------------------

        /// <summary>
        /// Generates the 1280x720 status card the bot streams into the meeting as its video tile.
        /// The heavy static artwork (background, avatar, titles) is rendered once and cached; each
        /// call only paints the dynamic status pill, waveform, timestamp and optional activity line.
        /// </summary>
        public static Bitmap CreateBotStatusCard(string botName, string statusMessage, string meetingId, bool isMuted = false, int tick = 0, string activityLine = null)
        {
            Bitmap background = GetCardBackground(botName ?? "Teams AI Assistant");
            var bmp = (Bitmap)background.Clone();

            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                // REC pulse
                bool pulse = (tick % 2 == 0);
                Color recColor = pulse ? Color.FromArgb(239, 68, 68) : Color.FromArgb(185, 28, 28);
                using (var recBrush = new SolidBrush(recColor))
                {
                    g.FillEllipse(recBrush, 1070, 95, 16, 16);
                }

                // Status pill
                int pillX = 365;
                int pillY = 240;
                int pillW = isMuted ? 140 : 260;
                int pillH = 38;
                Color pillBg = isMuted ? Color.FromArgb(127, 29, 29) : Color.FromArgb(6, 78, 59);
                Color pillBorder = isMuted ? Color.FromArgb(239, 68, 68) : Color.FromArgb(16, 185, 129);
                Color pillText = isMuted ? Color.FromArgb(254, 202, 202) : Color.FromArgb(167, 243, 208);
                string pillLabel = isMuted ? "● MUTED" : "● ACTIVE & LISTENING";

                using (var pillBrush = new SolidBrush(pillBg))
                using (var pillPen = new Pen(pillBorder, 1.5f))
                {
                    g.FillRectangle(pillBrush, pillX, pillY, pillW, pillH);
                    g.DrawRectangle(pillPen, pillX, pillY, pillW, pillH);
                }

                using (var pillFont = new Font("Segoe UI", 13, FontStyle.Bold))
                using (var labelBrush = new SolidBrush(pillText))
                {
                    g.DrawString(pillLabel, pillFont, labelBrush, new PointF(pillX + 15, pillY + 8));
                }

                // Waveform bars
                int barStartX = 650;
                int barY = 240;
                using (var barBrush = new SolidBrush(Color.FromArgb(56, 189, 248)))
                {
                    for (int i = 0; i < 16; i++)
                    {
                        double wave = Math.Abs(Math.Sin((tick * 0.3) + (i * 0.5)));
                        int barH = isMuted ? 4 : (int)(wave * 30) + 6;
                        int currentBarY = barY + (38 - barH) / 2;
                        g.FillRectangle(barBrush, barStartX + (i * 12), currentBarY, 7, barH);
                    }
                }

                using (var fontDetails = new Font("Segoe UI", 15, FontStyle.Regular))
                using (var fontDetailsBold = new Font("Segoe UI", 15, FontStyle.Bold))
                using (var labelBrush = new SolidBrush(Color.FromArgb(100, 116, 139)))
                using (var valBrush = new SolidBrush(Color.FromArgb(226, 232, 240)))
                using (var accentBrush = new SolidBrush(Color.FromArgb(125, 211, 252)))
                {
                    g.DrawString("Call Session ID:", fontDetailsBold, labelBrush, new PointF(140, 400));
                    g.DrawString(meetingId ?? string.Empty, fontDetails, valBrush, new PointF(330, 400));

                    g.DrawString("Media Services:", fontDetailsBold, labelBrush, new PointF(140, 445));
                    g.DrawString("Per-speaker audio, screen-share video recording, minutes of meeting", fontDetails, valBrush, new PointF(330, 445));

                    g.DrawString("Operational State:", fontDetailsBold, labelBrush, new PointF(140, 490));
                    string opState = statusMessage ?? (isMuted ? "Audio output muted - still capturing" : "Recording audio and video");
                    g.DrawString(opState, fontDetails, valBrush, new PointF(330, 490));

                    g.DrawString("Local Timestamp:", fontDetailsBold, labelBrush, new PointF(140, 535));
                    g.DrawString(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), fontDetails, valBrush, new PointF(330, 535));

                    if (!string.IsNullOrWhiteSpace(activityLine))
                    {
                        g.DrawString("Activity:", fontDetailsBold, labelBrush, new PointF(140, 575));
                        g.DrawString(activityLine, fontDetails, accentBrush, new PointF(330, 575));
                    }
                }
            }

            return bmp;
        }

        /// <summary>
        /// Renders a 1280x720 frame that shows an arbitrary content image (chart, TDA answer card,
        /// snapshot, etc.) with a thin title bar, for the bot's outgoing video tile. This is the
        /// "share screen to show visualisation" capability (option a - the bot's video tile displays
        /// content, not a real Teams screen-share) - see BOT_CAPABILITY_EXPECTATIONS.md section 4.
        /// Off by default; CallHandler only calls this while a visualization is explicitly active,
        /// and reverts to <see cref="CreateBotStatusCard"/> otherwise.
        /// </summary>
        public static Bitmap CreateVisualizationFrame(Bitmap contentImage, string title)
        {
            var bmp = new Bitmap(1280, 720, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                using (var bgBrush = new SolidBrush(Color.FromArgb(15, 23, 42)))
                {
                    g.FillRectangle(bgBrush, 0, 0, 1280, 720);
                }

                const int titleBarHeight = 56;
                using (var titleBrush = new SolidBrush(Color.FromArgb(24, 32, 50)))
                {
                    g.FillRectangle(titleBrush, 0, 0, 1280, titleBarHeight);
                }

                using (var titleFont = new Font("Segoe UI", 18, FontStyle.Bold))
                using (var titleTextBrush = new SolidBrush(Color.White))
                {
                    g.DrawString(title ?? "Teams AI Assistant - Visualization", titleFont, titleTextBrush, new PointF(24, 12));
                }

                if (contentImage != null)
                {
                    int areaX = 20;
                    int areaY = titleBarHeight + 20;
                    int areaW = 1280 - (areaX * 2);
                    int areaH = 720 - areaY - 20;

                    float scale = Math.Min((float)areaW / contentImage.Width, (float)areaH / contentImage.Height);
                    int drawW = Math.Max(1, (int)(contentImage.Width * scale));
                    int drawH = Math.Max(1, (int)(contentImage.Height * scale));
                    int drawX = areaX + ((areaW - drawW) / 2);
                    int drawY = areaY + ((areaH - drawH) / 2);

                    using (var frameBrush = new SolidBrush(Color.White))
                    {
                        g.FillRectangle(frameBrush, drawX - 4, drawY - 4, drawW + 8, drawH + 8);
                    }

                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(contentImage, new Rectangle(drawX, drawY, drawW, drawH));
                }
            }

            return bmp;
        }

        private static Bitmap GetCardBackground(string botName)
        {
            lock (CardLock)
            {
                if (cachedCardBackground != null && cachedCardBotName == botName)
                {
                    return cachedCardBackground;
                }

                cachedCardBackground?.Dispose();
                cachedCardBackground = RenderCardBackground(botName);
                cachedCardBotName = botName;
                return cachedCardBackground;
            }
        }

        private static Bitmap RenderCardBackground(string botName)
        {
            var bmp = new Bitmap(1280, 720, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                using (var bgBrush = new LinearGradientBrush(new Rectangle(0, 0, 1280, 720), Color.FromArgb(15, 23, 42), Color.FromArgb(30, 41, 59), 45f))
                {
                    g.FillRectangle(bgBrush, 0, 0, 1280, 720);
                }

                using (var cardBrush = new SolidBrush(Color.FromArgb(24, 32, 50)))
                using (var cardBorderPen = new Pen(Color.FromArgb(51, 65, 85), 2f))
                {
                    var cardRect = new Rectangle(80, 60, 1120, 600);
                    g.FillRectangle(cardBrush, cardRect);
                    g.DrawRectangle(cardBorderPen, cardRect);
                }

                using (var recFont = new Font("Segoe UI", 13, FontStyle.Bold))
                using (var recTextBrush = new SolidBrush(Color.FromArgb(241, 245, 249)))
                {
                    g.DrawString("REC", recFont, recTextBrush, new PointF(1095, 92));
                }

                // Avatar
                int iconX = 130;
                int iconY = 140;
                int iconSize = 190;

                using (var glowPen = new Pen(Color.FromArgb(70, 56, 189, 248), 6f))
                {
                    g.DrawEllipse(glowPen, iconX - 5, iconY - 5, iconSize + 10, iconSize + 10);
                }

                using (var circleBrush = new LinearGradientBrush(new Rectangle(iconX, iconY, iconSize, iconSize), Color.FromArgb(14, 116, 144), Color.FromArgb(30, 58, 138), 60f))
                {
                    g.FillEllipse(circleBrush, iconX, iconY, iconSize, iconSize);
                }

                using (var borderPen = new Pen(Color.FromArgb(56, 189, 248), 3f))
                {
                    g.DrawEllipse(borderPen, iconX, iconY, iconSize, iconSize);
                }

                int headX = iconX + 45;
                int headY = iconY + 50;
                int headW = 100;
                int headH = 80;

                using (var headBrush = new SolidBrush(Color.FromArgb(248, 250, 252)))
                {
                    g.FillPie(headBrush, headX, headY - 10, headW, headH + 20, 0, 360);
                }

                using (var antennaPen = new Pen(Color.FromArgb(56, 189, 248), 4f))
                {
                    g.DrawLine(antennaPen, iconX + 95, iconY + 28, iconX + 95, iconY + 45);
                }

                using (var tipBrush = new SolidBrush(Color.FromArgb(56, 189, 248)))
                {
                    g.FillEllipse(tipBrush, iconX + 90, iconY + 20, 10, 10);
                }

                using (var visorBrush = new SolidBrush(Color.FromArgb(15, 23, 42)))
                {
                    g.FillRectangle(visorBrush, headX + 12, headY + 22, 76, 28);
                }

                using (var eyeBrush = new SolidBrush(Color.FromArgb(56, 189, 248)))
                {
                    g.FillEllipse(eyeBrush, headX + 22, headY + 26, 18, 18);
                    g.FillEllipse(eyeBrush, headX + 60, headY + 26, 18, 18);
                }

                using (var reflectBrush = new SolidBrush(Color.White))
                {
                    g.FillEllipse(reflectBrush, headX + 26, headY + 29, 6, 6);
                    g.FillEllipse(reflectBrush, headX + 64, headY + 29, 6, 6);
                }

                using (var smilePen = new Pen(Color.FromArgb(56, 189, 248), 3f))
                {
                    g.DrawArc(smilePen, headX + 32, headY + 54, 36, 16, 20, 140);
                }

                using (var fontTitle = new Font("Segoe UI", 34, FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString(botName, fontTitle, textBrush, new PointF(360, 140));
                }

                using (var fontSub = new Font("Segoe UI", 16, FontStyle.Regular))
                using (var subBrush = new SolidBrush(Color.FromArgb(148, 163, 184)))
                {
                    g.DrawString("Real-Time Meeting Intelligence & Media Bot", fontSub, subBrush, new PointF(365, 195));
                }

                using (var divPen = new Pen(Color.FromArgb(51, 65, 85), 1.5f))
                {
                    g.DrawLine(divPen, 130, 370, 1140, 370);
                }

                using (var fontFooter = new Font("Segoe UI", 12, FontStyle.Italic))
                using (var footerBrush = new SolidBrush(Color.FromArgb(71, 85, 105)))
                {
                    g.DrawString("Secure Microsoft Teams Bot Platform • Media Engine Online", fontFooter, footerBrush, new PointF(140, 625));
                }
            }

            return bmp;
        }
    }
}
