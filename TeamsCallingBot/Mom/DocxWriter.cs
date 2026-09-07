namespace TeamsCallingBot.Mom
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;
    using System.Security;
    using System.Text;

    /// <summary>
    /// Dependency-free Word (.docx / OOXML) writer. A .docx is a ZIP with a handful of XML parts;
    /// this class emits exactly the parts Word needs: [Content_Types].xml, _rels/.rels,
    /// word/document.xml, word/styles.xml, word/_rels/document.xml.rels and word/media/*.jpg.
    /// Supports headings, paragraphs (bold/italic), bullet lists, simple tables and inline images.
    /// Kept deliberately small - it is not a general purpose Word library.
    /// </summary>
    public sealed class DocxWriter
    {
        private const int EmuPerInch = 914400;

        private readonly StringBuilder body = new StringBuilder();
        private readonly List<ImagePart> images = new List<ImagePart>();
        private int drawingId = 1;

        public void AddTitle(string text)
        {
            this.body.Append("<w:p><w:pPr><w:pStyle w:val=\"Title\"/></w:pPr>").Append(Run(text)).Append("</w:p>");
        }

        public void AddHeading(string text, int level = 1)
        {
            level = Math.Max(1, Math.Min(3, level));
            this.body.Append($"<w:p><w:pPr><w:pStyle w:val=\"Heading{level}\"/></w:pPr>").Append(Run(text)).Append("</w:p>");
        }

        public void AddParagraph(string text, bool bold = false, bool italic = false, string color = null, int? sizeHalfPoints = null)
        {
            if (text == null)
            {
                text = string.Empty;
            }

            this.body.Append("<w:p>");
            foreach (var line in text.Replace("\r\n", "\n").Split('\n'))
            {
                this.body.Append(Run(line, bold, italic, color, sizeHalfPoints));
                this.body.Append("<w:r><w:br/></w:r>");
            }

            // Remove the trailing line break we just appended.
            this.body.Length -= "<w:r><w:br/></w:r>".Length;
            this.body.Append("</w:p>");
        }

        public void AddLabelValue(string label, string value)
        {
            this.body.Append("<w:p>")
                .Append(Run(label + ": ", bold: true))
                .Append(Run(value ?? string.Empty))
                .Append("</w:p>");
        }

        public void AddBullets(IEnumerable<string> items)
        {
            if (items == null)
            {
                return;
            }

            foreach (var item in items)
            {
                this.body.Append("<w:p><w:pPr><w:ind w:left=\"540\" w:hanging=\"270\"/></w:pPr>")
                    .Append(Run("•  " + (item ?? string.Empty)))
                    .Append("</w:p>");
            }
        }

        public void AddNumbered(IEnumerable<string> items)
        {
            if (items == null)
            {
                return;
            }

            int i = 1;
            foreach (var item in items)
            {
                this.body.Append("<w:p><w:pPr><w:ind w:left=\"540\" w:hanging=\"360\"/></w:pPr>")
                    .Append(Run($"{i}.  " + (item ?? string.Empty)))
                    .Append("</w:p>");
                i++;
            }
        }

        public void AddTable(IList<string> headers, IEnumerable<IList<string>> rows)
        {
            this.body.Append("<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/><w:tblW w:w=\"5000\" w:type=\"pct\"/>")
                .Append("<w:tblBorders>")
                .Append("<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"9CA3AF\"/>")
                .Append("<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"9CA3AF\"/>")
                .Append("<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"9CA3AF\"/>")
                .Append("<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"9CA3AF\"/>")
                .Append("<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D1D5DB\"/>")
                .Append("<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D1D5DB\"/>")
                .Append("</w:tblBorders></w:tblPr>");

            if (headers != null && headers.Count > 0)
            {
                this.body.Append("<w:tr>");
                foreach (var h in headers)
                {
                    this.body.Append("<w:tc><w:tcPr><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"1F2937\"/></w:tcPr><w:p>")
                        .Append(Run(h ?? string.Empty, bold: true, color: "FFFFFF"))
                        .Append("</w:p></w:tc>");
                }

                this.body.Append("</w:tr>");
            }

            if (rows != null)
            {
                foreach (var row in rows)
                {
                    this.body.Append("<w:tr>");
                    foreach (var cell in row)
                    {
                        this.body.Append("<w:tc><w:p>").Append(Run(cell ?? string.Empty)).Append("</w:p></w:tc>");
                    }

                    this.body.Append("</w:tr>");
                }
            }

            this.body.Append("</w:tbl>");
            this.AddParagraph(string.Empty);
        }

        /// <summary>Embeds a JPEG image scaled to the given width (inches), preserving aspect ratio.</summary>
        public bool AddImage(string jpegPath, double widthInches = 6.0, string caption = null)
        {
            try
            {
                if (!File.Exists(jpegPath))
                {
                    return false;
                }

                byte[] bytes = File.ReadAllBytes(jpegPath);
                int pxW;
                int pxH;
                using (var ms = new MemoryStream(bytes))
                using (var img = Image.FromStream(ms, false, false))
                {
                    pxW = img.Width;
                    pxH = img.Height;
                }

                if (pxW <= 0 || pxH <= 0)
                {
                    return false;
                }

                long cx = (long)(widthInches * EmuPerInch);
                long cy = (long)(cx * ((double)pxH / pxW));

                int imageIndex = this.images.Count + 1;
                string rid = $"rIdImg{imageIndex}";
                this.images.Add(new ImagePart { RelationshipId = rid, FileName = $"image{imageIndex}.jpg", Data = bytes });

                int id = this.drawingId++;
                this.body.Append("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:drawing>")
                    .Append($"<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{cx}\" cy=\"{cy}\"/>")
                    .Append($"<wp:docPr id=\"{id}\" name=\"Picture {id}\"/>")
                    .Append("<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">")
                    .Append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">")
                    .Append("<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">")
                    .Append($"<pic:nvPicPr><pic:cNvPr id=\"{id}\" name=\"{Esc(Path.GetFileName(jpegPath))}\"/><pic:cNvPicPr/></pic:nvPicPr>")
                    .Append($"<pic:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>")
                    .Append($"<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>")
                    .Append("</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>");

                if (!string.IsNullOrWhiteSpace(caption))
                {
                    this.body.Append("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>")
                        .Append(Run(caption, italic: true, color: "6B7280", sizeHalfPoints: 18))
                        .Append("</w:p>");
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public void AddPageBreak()
        {
            this.body.Append("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>");
        }

        public void Save(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? ".");
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                WriteEntry(zip, "[Content_Types].xml", this.ContentTypesXml());
                WriteEntry(zip, "_rels/.rels", RootRelsXml);
                WriteEntry(zip, "word/document.xml", this.DocumentXml());
                WriteEntry(zip, "word/styles.xml", StylesXml);
                WriteEntry(zip, "word/_rels/document.xml.rels", this.DocumentRelsXml());
                foreach (var image in this.images)
                {
                    var entry = zip.CreateEntry("word/media/" + image.FileName, CompressionLevel.NoCompression);
                    using (var s = entry.Open())
                    {
                        s.Write(image.Data, 0, image.Data.Length);
                    }
                }
            }
        }

        // ------------------------------------------------------------------

        private static string Run(string text, bool bold = false, bool italic = false, string color = null, int? sizeHalfPoints = null)
        {
            var sb = new StringBuilder("<w:r>");
            if (bold || italic || color != null || sizeHalfPoints.HasValue)
            {
                sb.Append("<w:rPr>");
                if (bold) sb.Append("<w:b/>");
                if (italic) sb.Append("<w:i/>");
                if (color != null) sb.Append($"<w:color w:val=\"{color}\"/>");
                if (sizeHalfPoints.HasValue) sb.Append($"<w:sz w:val=\"{sizeHalfPoints.Value}\"/>");
                sb.Append("</w:rPr>");
            }

            sb.Append("<w:t xml:space=\"preserve\">").Append(Esc(text ?? string.Empty)).Append("</w:t></w:r>");
            return sb.ToString();
        }

        private static string Esc(string text)
        {
            // Strip characters that are illegal in XML 1.0 before escaping.
            var sb = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                if (c == 0x9 || c == 0xA || c == 0xD || (c >= 0x20 && c <= 0xD7FF) || (c >= 0xE000 && c <= 0xFFFD))
                {
                    sb.Append(c);
                }
            }

            return SecurityElement.Escape(sb.ToString());
        }

        private static void WriteEntry(ZipArchive zip, string name, string xml)
        {
            var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var s = entry.Open())
            {
                var bytes = new UTF8Encoding(false).GetBytes(xml);
                s.Write(bytes, 0, bytes.Length);
            }
        }

        private string DocumentXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                   "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                   "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                   "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                   "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                   "<w:body>" + this.body +
                   "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1080\" w:right=\"1080\" w:bottom=\"1080\" w:left=\"1080\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>" +
                   "</w:body></w:document>";
        }

        private string ContentTypesXml()
        {
            var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>");
            sb.Append("<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>");
            sb.Append("<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>");
            sb.Append("</Types>");
            return sb.ToString();
        }

        private string DocumentRelsXml()
        {
            var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            sb.Append("<Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            foreach (var image in this.images)
            {
                sb.Append($"<Relationship Id=\"{image.RelationshipId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{image.FileName}\"/>");
            }

            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private const string RootRelsXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "</Relationships>";

        private const string StylesXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
            "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\"/><w:sz w:val=\"22\"/><w:lang w:val=\"en-IN\"/></w:rPr></w:rPrDefault>" +
            "<w:pPrDefault><w:pPr><w:spacing w:after=\"120\" w:line=\"264\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>" +
            "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Title\"><w:name w:val=\"Title\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:after=\"240\"/></w:pPr><w:rPr><w:b/><w:color w:val=\"111827\"/><w:sz w:val=\"52\"/></w:rPr></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:spacing w:before=\"360\" w:after=\"120\"/><w:pBdr><w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"2563EB\"/></w:pBdr><w:outlineLvl w:val=\"0\"/></w:pPr><w:rPr><w:b/><w:color w:val=\"1E3A8A\"/><w:sz w:val=\"32\"/></w:rPr></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:spacing w:before=\"240\" w:after=\"80\"/><w:outlineLvl w:val=\"1\"/></w:pPr><w:rPr><w:b/><w:color w:val=\"1F2937\"/><w:sz w:val=\"26\"/></w:rPr></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"Heading3\"><w:name w:val=\"heading 3\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:spacing w:before=\"160\" w:after=\"60\"/><w:outlineLvl w:val=\"2\"/></w:pPr><w:rPr><w:b/><w:i/><w:color w:val=\"374151\"/><w:sz w:val=\"23\"/></w:rPr></w:style>" +
            "<w:style w:type=\"table\" w:styleId=\"TableGrid\"><w:name w:val=\"Table Grid\"/><w:tblPr><w:tblCellMar><w:left w:w=\"80\" w:type=\"dxa\"/><w:right w:w=\"80\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style>" +
            "</w:styles>";

        private sealed class ImagePart
        {
            public string RelationshipId;
            public string FileName;
            public byte[] Data;
        }
    }
}
