using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace Sadar.Cli
{
    /// <summary>
    /// בונה קובץ docx מינימלי לצורך בדיקות.
    ///
    /// לא חלק מהמוצר — כלי עזר שמאפשר לבדוק את הצינור על מסמך
    /// בעל מבנה ידוע מראש, בלי לתלות את הבדיקות בקובץ חיצוני.
    /// </summary>
    public static class DocxBuilder
    {
        /// <summary>קטע טקסט בתוך פסקה, עם עיצוב משלו.</summary>
        public sealed class Run
        {
            public string Text;
            public bool Bold;

            public static Run Of(string text, bool bold = false)
            {
                return new Run { Text = text, Bold = bold };
            }
        }

        public sealed class Para
        {
            public string Text;
            /// <summary>אם מלא — הפסקה נבנית מכמה הרצות במקום מטקסט אחד.</summary>
            public List<Run> Runs;
            public bool Bold;
            public double Size = 12;
            public string Align = "both";
            public string StyleId;

            public static Para Of(string text, bool bold = false, double size = 12,
                                  string align = "both", string styleId = null)
            {
                return new Para { Text = text, Bold = bold, Size = size, Align = align, StyleId = styleId };
            }

            /// <summary>פסקה שמורכבת מכמה הרצות — למשל מילה מודגשת בתוך משפט.</summary>
            public static Para OfRuns(double size, string align, params Run[] runs)
            {
                var p = new Para { Size = size, Align = align, Runs = new List<Run>(runs) };
                var sb = new StringBuilder();
                foreach (var r in runs) sb.Append(r.Text);
                p.Text = sb.ToString();
                return p;
            }
        }

        public static void Create(string path, IEnumerable<Para> paragraphs)
        {
            if (File.Exists(path)) File.Delete(path);

            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                Write(zip, "[Content_Types].xml", ContentTypes);
                Write(zip, "_rels/.rels", RootRels);
                Write(zip, "word/_rels/document.xml.rels", DocumentRels);
                Write(zip, "word/styles.xml", Styles);
                Write(zip, "word/document.xml", BuildDocument(paragraphs));
            }
        }

        private static void Write(ZipArchive zip, string name, string content)
        {
            var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var s = entry.Open())
            using (var w = new StreamWriter(s, new UTF8Encoding(false)))
                w.Write(content);
        }

        private static string BuildDocument(IEnumerable<Para> paragraphs)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
            sb.Append("<w:body>");

            foreach (var p in paragraphs)
            {
                int halfPoints = (int)Math.Round(p.Size * 2);

                sb.Append("<w:p><w:pPr>");
                if (!string.IsNullOrEmpty(p.StyleId))
                    sb.Append("<w:pStyle w:val=\"").Append(Esc(p.StyleId)).Append("\"/>");
                sb.Append("<w:bidi/>");
                sb.Append("<w:jc w:val=\"").Append(p.Align).Append("\"/>");
                sb.Append("</w:pPr>");

                if (p.Runs != null && p.Runs.Count > 0)
                {
                    foreach (var r in p.Runs) AppendRun(sb, r.Text, r.Bold, halfPoints);
                }
                else if (!string.IsNullOrEmpty(p.Text))
                {
                    AppendRun(sb, p.Text, p.Bold, halfPoints);
                }

                sb.Append("</w:p>");
            }

            sb.Append("<w:sectPr><w:bidi/></w:sectPr>");
            sb.Append("</w:body></w:document>");
            return sb.ToString();
        }

        private static void AppendRun(StringBuilder sb, string text, bool bold, int halfPoints)
        {
            if (string.IsNullOrEmpty(text)) return;
            sb.Append("<w:r><w:rPr>");
            if (bold) sb.Append("<w:b/><w:bCs/>");
            sb.Append("<w:sz w:val=\"").Append(halfPoints).Append("\"/>");
            sb.Append("<w:szCs w:val=\"").Append(halfPoints).Append("\"/>");
            sb.Append("<w:rtl/>");
            sb.Append("</w:rPr>");
            sb.Append("<w:t xml:space=\"preserve\">").Append(Esc(text)).Append("</w:t>");
            sb.Append("</w:r>");
        }

        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        private const string ContentTypes =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
            "</Types>";

        private const string RootRels =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "</Relationships>";

        private const string DocumentRels =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
            "</Relationships>";

        private const string Styles =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
            "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">" +
            "<w:name w:val=\"Normal\"/><w:qFormat/>" +
            "</w:style>" +
            "</w:styles>";
    }
}
