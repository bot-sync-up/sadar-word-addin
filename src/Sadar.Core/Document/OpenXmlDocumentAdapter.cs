using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Sadar.Core.Model;

namespace Sadar.Core.Document
{
    /// <summary>
    /// עבודה ישירה על קובץ docx, בלי וורד ובלי ספריות חיצוניות.
    ///
    /// משמש את כלי שורת הפקודה ואת מערך הבדיקות: כך אפשר לבדוק את כל
    /// הצינור — סריקה, החלטה, החלה ואימות — על מחשב שאין בו וורד כלל.
    /// </summary>
    public sealed class OpenXmlDocumentAdapter : IDocumentAdapter
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private readonly string _path;
        private XDocument _document;
        private XDocument _styles;
        private List<XElement> _paragraphs;
        private byte[] _batchSnapshotDocument;
        private byte[] _batchSnapshotStyles;

        /// <summary>מיפוי שם מוצג -> styleId עבור סגנונות מובנים.</summary>
        private static readonly Dictionary<string, string> BuiltInStyleIds =
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "כותרת 1", "Heading1" },
                { "כותרת 2", "Heading2" },
                { "כותרת 3", "Heading3" },
                { "כותרת 4", "Heading4" },
                { "Heading 1", "Heading1" },
                { "Heading 2", "Heading2" },
                { "Heading 3", "Heading3" },
                { "Heading 4", "Heading4" },
                { "רגיל", "Normal" },
                { "Normal", "Normal" },
            };

        private static readonly Dictionary<string, string> BuiltInDisplayNames =
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Heading1", "כותרת 1" },
                { "Heading2", "כותרת 2" },
                { "Heading3", "כותרת 3" },
                { "Heading4", "כותרת 4" },
                { "Normal", "רגיל" },
            };

        public OpenXmlDocumentAdapter(string path)
        {
            _path = path;
            Load();
        }

        private void Load()
        {
            using (var zip = ZipFile.Open(_path, ZipArchiveMode.Read))
            {
                _document = ReadXml(zip, "word/document.xml");
                if (_document == null)
                    throw new InvalidDataException("הקובץ אינו מסמך וורד תקין: לא נמצא word/document.xml");

                _styles = ReadXml(zip, "word/styles.xml");
            }
            IndexParagraphs();
        }

        private static XDocument ReadXml(ZipArchive zip, string entryName)
        {
            var entry = zip.GetEntry(entryName);
            if (entry == null) return null;
            using (var s = entry.Open())
                return XDocument.Load(s, LoadOptions.PreserveWhitespace);
        }

        private void IndexParagraphs()
        {
            var body = _document.Root.Element(W + "body");
            if (body == null)
                throw new InvalidDataException("המסמך אינו מכיל body.");

            // רק פסקאות ברמת הגוף. פסקאות בתוך טבלאות נסרקות אך אינן ניתנות למחיקה.
            _paragraphs = body.Descendants(W + "p").ToList();
        }

        public int ParagraphCount { get { return _paragraphs.Count; } }

        // ---------- קריאה ----------

        public List<ParagraphInfo> ReadAll()
        {
            var list = new List<ParagraphInfo>(_paragraphs.Count);
            for (int i = 0; i < _paragraphs.Count; i++)
                list.Add(ReadParagraph(_paragraphs[i], i));
            return list;
        }

        private ParagraphInfo ReadParagraph(XElement p, int index)
        {
            var info = new ParagraphInfo
            {
                Index = index,
                Text = ExtractText(p),
                StyleName = "רגיל",
                Alignment = Align.Unknown
            };

            var pPr = p.Element(W + "pPr");
            if (pPr != null)
            {
                var pStyle = pPr.Element(W + "pStyle");
                if (pStyle != null)
                {
                    string id = Attr(pStyle, "val");
                    info.StyleName = DisplayNameForStyleId(id);
                }

                var jc = pPr.Element(W + "jc");
                if (jc != null) info.Alignment = ParseAlign(Attr(jc, "val"));

                var outline = pPr.Element(W + "outlineLvl");
                if (outline != null)
                {
                    int lvl;
                    if (int.TryParse(Attr(outline, "val"), out lvl)) info.OutlineLevel = lvl;
                }

                if (pPr.Element(W + "numPr") != null) info.IsListItem = true;

                var rPr = pPr.Element(W + "rPr");
                if (rPr != null) ReadRunProps(rPr, info);
            }

            // מאפייני ההרצה הראשונה שיש בה טקסט — היא הקובעת את מראה הפסקה
            foreach (var run in p.Elements(W + "r"))
            {
                var t = run.Element(W + "t");
                if (t == null || string.IsNullOrEmpty(t.Value)) continue;
                var rPr = run.Element(W + "rPr");
                if (rPr != null) ReadRunProps(rPr, info);
                break;
            }

            if (info.Alignment == Align.Unknown) info.Alignment = Align.Right;
            return info;
        }

        /// <summary>
        /// בעברית המאפיינים הקובעים הם של הכתב המורכב (bCs, szCs) ולא b/sz.
        /// קריאה מהשדות הלטיניים בלבד מחמיצה כותרות מודגשות בעברית.
        /// </summary>
        private static void ReadRunProps(XElement rPr, ParagraphInfo info)
        {
            var bCs = rPr.Element(W + "bCs");
            var b = rPr.Element(W + "b");
            if (bCs != null) info.Bold = IsOn(bCs);
            else if (b != null) info.Bold = IsOn(b);

            var szCs = rPr.Element(W + "szCs");
            var sz = rPr.Element(W + "sz");
            var chosen = szCs ?? sz;
            if (chosen != null)
            {
                double halfPoints;
                if (double.TryParse(Attr(chosen, "val"),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out halfPoints))
                    info.FontSize = halfPoints / 2.0;
            }
        }

        private static bool IsOn(XElement toggle)
        {
            string v = Attr(toggle, "val");
            if (v == null) return true; // נוכחות האלמנט בלי val משמעה דלוק
            return v != "0" && !string.Equals(v, "false", StringComparison.OrdinalIgnoreCase);
        }

        private static string ExtractText(XElement p)
        {
            var sb = new StringBuilder();
            foreach (var node in p.Descendants())
            {
                if (node.Name == W + "t") sb.Append(node.Value);
                else if (node.Name == W + "tab") sb.Append('\t');
                else if (node.Name == W + "br") sb.Append(' ');
            }
            return sb.ToString();
        }

        private static string Attr(XElement e, string name)
        {
            var a = e.Attribute(W + name);
            return a == null ? null : a.Value;
        }

        private static Align ParseAlign(string val)
        {
            switch (val)
            {
                case "right": return Align.Right;
                case "left": return Align.Left;
                case "center": return Align.Center;
                case "both":
                case "distribute": return Align.Justify;
                default: return Align.Unknown;
            }
        }

        // ---------- סגנונות ----------

        public List<string> GetStyleNames()
        {
            var names = new List<string>();
            if (_styles == null) return names;

            foreach (var style in _styles.Root.Elements(W + "style"))
            {
                var nameEl = style.Element(W + "name");
                string id = Attr(style, "styleId");
                string display = nameEl != null ? Attr(nameEl, "val") : null;
                names.Add(DisplayNameForStyleId(id) ?? display ?? id);
            }
            return names;
        }

        private string DisplayNameForStyleId(string id)
        {
            if (string.IsNullOrEmpty(id)) return "רגיל";

            string display;
            if (BuiltInDisplayNames.TryGetValue(id, out display)) return display;

            if (_styles != null)
            {
                foreach (var style in _styles.Root.Elements(W + "style"))
                {
                    if (!string.Equals(Attr(style, "styleId"), id, StringComparison.Ordinal)) continue;
                    var nameEl = style.Element(W + "name");
                    if (nameEl != null)
                    {
                        string val = Attr(nameEl, "val");
                        // "heading 1" -> "כותרת 1"
                        if (!string.IsNullOrEmpty(val))
                        {
                            string mapped;
                            if (BuiltInStyleIds.TryGetValue(val, out mapped) &&
                                BuiltInDisplayNames.TryGetValue(mapped, out display))
                                return display;
                            return val;
                        }
                    }
                }
            }
            return id;
        }

        private static string StyleIdFor(string displayName)
        {
            string id;
            if (BuiltInStyleIds.TryGetValue(displayName, out id)) return id;
            // סגנון מותאם — מזהה בלי רווחים
            return displayName.Replace(" ", "");
        }

        public bool EnsureStyleExists(string styleName)
        {
            if (_styles == null) return false;

            string id = StyleIdFor(styleName);
            foreach (var style in _styles.Root.Elements(W + "style"))
                if (string.Equals(Attr(style, "styleId"), id, StringComparison.Ordinal))
                    return true;

            if (string.Equals(id, "Normal", StringComparison.Ordinal)) return true;

            int level = 0;
            if (id.StartsWith("Heading", StringComparison.Ordinal))
                int.TryParse(id.Substring("Heading".Length), out level);

            var newStyle = new XElement(W + "style",
                new XAttribute(W + "type", "paragraph"),
                new XAttribute(W + "styleId", id),
                new XElement(W + "name", new XAttribute(W + "val",
                    level > 0 ? "heading " + level : styleName)),
                new XElement(W + "basedOn", new XAttribute(W + "val", "Normal")),
                new XElement(W + "qFormat"),
                new XElement(W + "pPr",
                    new XElement(W + "keepNext"),
                    new XElement(W + "bidi"),
                    new XElement(W + "spacing",
                        new XAttribute(W + "before", "240"),
                        new XAttribute(W + "after", "120")),
                    new XElement(W + "jc", new XAttribute(W + "val", "center")),
                    level > 0
                        ? new XElement(W + "outlineLvl", new XAttribute(W + "val", (level - 1).ToString()))
                        : null),
                new XElement(W + "rPr",
                    new XElement(W + "b"),
                    new XElement(W + "bCs"),
                    new XElement(W + "sz", new XAttribute(W + "val", HeadingHalfPoints(level))),
                    new XElement(W + "szCs", new XAttribute(W + "val", HeadingHalfPoints(level)))));

            _styles.Root.Add(newStyle);
            return true;
        }

        private static string HeadingHalfPoints(int level)
        {
            switch (level)
            {
                case 1: return "32"; // 16pt
                case 2: return "28"; // 14pt
                case 3: return "26"; // 13pt
                default: return "24";
            }
        }

        public void ApplyStyle(int paragraphIndex, string styleName)
        {
            var p = GetParagraph(paragraphIndex);
            var pPr = p.Element(W + "pPr");
            if (pPr == null)
            {
                pPr = new XElement(W + "pPr");
                p.AddFirst(pPr);
            }

            var pStyle = pPr.Element(W + "pStyle");
            string id = StyleIdFor(styleName);

            if (pStyle == null)
            {
                // pStyle חייב להיות ראשון בתוך pPr לפי הסכמה
                pPr.AddFirst(new XElement(W + "pStyle", new XAttribute(W + "val", id)));
            }
            else
            {
                pStyle.SetAttributeValue(W + "val", id);
            }
        }

        // ---------- כתיבת טקסט (ניקוי רווחים בלבד) ----------

        /// <summary>
        /// כותב את הטקסט המנוקה חזרה לפסקה תוך שמירה מלאה על חלוקת ההרצות.
        ///
        /// הניקוי מסיר תווי רווח בלבד ולעולם אינו מוסיף או משנה תו אחר,
        /// ולכן אפשר להתאים את הטקסט החדש להרצות הקיימות בהליכה אחת:
        /// כל תו שמופיע בשניהם נשמר בהרצה שלו, וכל תו שהוסר פשוט מדולג.
        /// כך הדגשות ועיצוב בתוך הפסקה נשארים במקומם.
        /// </summary>
        public void SetParagraphText(int paragraphIndex, string normalizedText)
        {
            var p = GetParagraph(paragraphIndex);
            var textNodes = p.Descendants(W + "t").ToList();
            if (textNodes.Count == 0) return;

            int t = 0;
            string target = normalizedText ?? string.Empty;

            for (int n = 0; n < textNodes.Count; n++)
            {
                string source = textNodes[n].Value;
                var sb = new StringBuilder(source.Length);

                foreach (char c in source)
                {
                    if (t >= target.Length) break;

                    if (target[t] == c)
                    {
                        sb.Append(c);
                        t++;
                    }
                    else if (IsWhite(c) && IsWhite(target[t]))
                    {
                        sb.Append(target[t]);
                        t++;
                    }
                    // אחרת: התו הוסר בניקוי — מדלגים עליו
                }

                // השארית נכתבת להרצה האחרונה, כדי שלא ייעלם דבר
                if (n == textNodes.Count - 1 && t < target.Length)
                {
                    sb.Append(target, t, target.Length - t);
                    t = target.Length;
                }

                SetTextValue(textNodes[n], sb.ToString());
            }
        }

        private static bool IsWhite(char c)
        {
            return c == ' ' || c == '\t' || c == ' ';
        }

        private static void SetTextValue(XElement t, string value)
        {
            t.Value = value;
            // רווח בקצה נשמר רק עם xml:space="preserve"
            if (value.Length > 0 && (value[0] == ' ' || value[value.Length - 1] == ' '))
                t.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            else
                t.SetAttributeValue(XNamespace.Xml + "space", null);
        }

        // ---------- מחיקה ----------

        public void DeleteParagraphs(IEnumerable<int> paragraphIndexes)
        {
            var sorted = new List<int>(paragraphIndexes);
            sorted.Sort();
            sorted.Reverse();

            foreach (int i in sorted)
            {
                if (i < 0 || i >= _paragraphs.Count) continue;
                var p = _paragraphs[i];
                // פסקה אחרונה בתא טבלה או במסמך חייבת להישאר
                if (p.Parent != null && p.Parent.Elements(W + "p").Count() <= 1) continue;
                p.Remove();
            }

            IndexParagraphs();
        }

        private XElement GetParagraph(int index)
        {
            if (index < 0 || index >= _paragraphs.Count)
                throw new ArgumentOutOfRangeException("index", "אינדקס פסקה מחוץ לטווח: " + index);
            return _paragraphs[index];
        }

        // ---------- אצווה וביטול ----------

        public void BeginBatch(string undoName)
        {
            _batchSnapshotDocument = Encoding.UTF8.GetBytes(_document.ToString(SaveOptions.DisableFormatting));
            _batchSnapshotStyles = _styles == null
                ? null
                : Encoding.UTF8.GetBytes(_styles.ToString(SaveOptions.DisableFormatting));
        }

        public void EndBatch()
        {
            _batchSnapshotDocument = null;
            _batchSnapshotStyles = null;
        }

        public void RollbackBatch()
        {
            if (_batchSnapshotDocument == null) return;

            _document = XDocument.Parse(
                Encoding.UTF8.GetString(_batchSnapshotDocument), LoadOptions.PreserveWhitespace);

            if (_batchSnapshotStyles != null)
                _styles = XDocument.Parse(
                    Encoding.UTF8.GetString(_batchSnapshotStyles), LoadOptions.PreserveWhitespace);

            IndexParagraphs();
            EndBatch();
        }

        // ---------- שמירה ----------

        public void Save()
        {
            SaveAs(_path);
        }

        public void SaveAs(string targetPath)
        {
            if (!string.Equals(targetPath, _path, StringComparison.OrdinalIgnoreCase))
                File.Copy(_path, targetPath, true);

            using (var zip = ZipFile.Open(targetPath, ZipArchiveMode.Update))
            {
                WriteXml(zip, "word/document.xml", _document);
                if (_styles != null) WriteXml(zip, "word/styles.xml", _styles);
            }
        }

        private static void WriteXml(ZipArchive zip, string entryName, XDocument doc)
        {
            var entry = zip.GetEntry(entryName);
            if (entry != null) entry.Delete();
            entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);

            using (var s = entry.Open())
            using (var w = new StreamWriter(s, new UTF8Encoding(false)))
            {
                w.Write(doc.Declaration != null
                    ? doc.Declaration.ToString()
                    : "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                w.Write("\r\n");
                w.Write(doc.ToString(SaveOptions.DisableFormatting));
            }
        }

        public void Dispose() { }
    }
}
