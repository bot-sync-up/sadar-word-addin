using System;
using System.Collections.Generic;
using Sadar.Core.Document;
using Sadar.Core.Model;
using Word = Microsoft.Office.Interop.Word;

namespace Sadar.Addin
{
    /// <summary>
    /// מימוש מעל המסמך הפתוח בוורד.
    ///
    /// כל הפעולות עוברות דרך אובייקט המסמך של וורד עצמו, ולכן הן
    /// מתועדות ב-Undo, מכובדות על ידי "עקוב אחר שינויים", ומשתלבות
    /// בכל מה שהמשתמש כבר עושה במסמך.
    /// </summary>
    public sealed class WordDocumentAdapter : IDocumentAdapter
    {
        private readonly Word.Application _app;
        private readonly Word.Document _doc;
        private bool _recordOpen;
        private bool _savedScreenUpdating;

        public WordDocumentAdapter(Word.Application app, Word.Document doc)
        {
            _app = app;
            _doc = doc;
        }

        public int ParagraphCount { get { return _doc.Paragraphs.Count; } }

        // ---------- קריאה ----------

        public List<ParagraphInfo> ReadAll()
        {
            var list = new List<ParagraphInfo>(_doc.Paragraphs.Count);

            bool prev = _app.ScreenUpdating;
            _app.ScreenUpdating = false;
            try
            {
                int index = 0;
                foreach (Word.Paragraph para in _doc.Paragraphs)
                {
                    list.Add(ReadParagraph(para, index));
                    index++;
                }
            }
            finally
            {
                _app.ScreenUpdating = prev;
            }

            return list;
        }

        private ParagraphInfo ReadParagraph(Word.Paragraph para, int index)
        {
            var range = para.Range;

            var info = new ParagraphInfo
            {
                Index = index,
                Text = StripParagraphMark(range.Text),
                OutlineLevel = (int)para.OutlineLevel
            };

            try
            {
                var style = para.get_Style() as Word.Style;
                info.StyleName = style != null ? style.NameLocal : "רגיל";
            }
            catch { info.StyleName = "רגיל"; }

            // בעברית המאפיין הקובע הוא הדו-כיווני (BoldBi/SizeBi) ולא Bold/Size.
            // קריאה מהמאפיין הלטיני בלבד מחמיצה כותרות מודגשות בעברית.
            try
            {
                object boldBi = range.BoldBi;
                int bi = Convert.ToInt32(boldBi);
                if (bi == 9999999) info.Bold = false;      // ערבוב בתוך הפסקה
                else if (bi != 0) info.Bold = true;
                else
                {
                    int b = Convert.ToInt32(range.Bold);
                    info.Bold = b != 0 && b != 9999999;
                }
            }
            catch { info.Bold = false; }

            try
            {
                float sizeBi = range.Font.SizeBi;
                float size = range.Font.Size;
                float chosen = sizeBi > 0 && sizeBi < 9999998 ? sizeBi : size;
                if (chosen > 0 && chosen < 9999998) info.FontSize = chosen;
            }
            catch { }

            try { info.Alignment = MapAlign(para.Format.Alignment); }
            catch { info.Alignment = Align.Unknown; }

            try
            {
                info.IsListItem = para.Range.ListFormat.ListType != Word.WdListType.wdListNoNumbering;
            }
            catch { }

            return info;
        }

        private static string StripParagraphMark(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            // וורד מחזיר את סימן סוף הפסקה כחלק מהטקסט
            return text.TrimEnd('\r', '\a', '\n');
        }

        private static Align MapAlign(Word.WdParagraphAlignment a)
        {
            switch (a)
            {
                case Word.WdParagraphAlignment.wdAlignParagraphRight: return Align.Right;
                case Word.WdParagraphAlignment.wdAlignParagraphLeft: return Align.Left;
                case Word.WdParagraphAlignment.wdAlignParagraphCenter: return Align.Center;
                case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                case Word.WdParagraphAlignment.wdAlignParagraphJustifyHi:
                case Word.WdParagraphAlignment.wdAlignParagraphJustifyLow:
                case Word.WdParagraphAlignment.wdAlignParagraphJustifyMed: return Align.Justify;
                default: return Align.Unknown;
            }
        }

        // ---------- סגנונות ----------

        public List<string> GetStyleNames()
        {
            var names = new List<string>();
            foreach (Word.Style s in _doc.Styles)
            {
                try { names.Add(s.NameLocal); } catch { }
            }
            return names;
        }

        public bool EnsureStyleExists(string styleName)
        {
            if (string.IsNullOrEmpty(styleName)) return false;

            try
            {
                object name = styleName;
                var existing = _doc.Styles.get_Item(ref name);
                return existing != null;
            }
            catch { /* לא קיים — ננסה ליצור */ }

            // סגנונות הכותרת המובנים קיימים תמיד בוורד, גם אם לא נעשה בהם שימוש.
            // אם הגענו לכאן מדובר בסגנון מותאם, ולכן ניצור אותו.
            try
            {
                object type = Word.WdStyleType.wdStyleTypeParagraph;
                var style = _doc.Styles.Add(styleName, ref type);
                style.set_BaseStyle("רגיל");
                style.Font.Bold = -1;
                style.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                style.ParagraphFormat.KeepWithNext = -1;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void ApplyStyle(int paragraphIndex, string styleName)
        {
            var para = GetParagraph(paragraphIndex);
            object style = styleName;
            para.set_Style(ref style);
        }

        // ---------- ניקוי רווחים ----------

        /// <summary>
        /// מסיר את תווי הרווח שהניקוי החליט להסיר, תו-תו, במקום לכתוב
        /// את הפסקה מחדש. כתיבה מחדש של Range.Text הורסת הדגשות בתוך הפסקה
        /// ומאבדת עיצוב פנימי — ולכן היא אסורה כאן.
        /// </summary>
        public void SetParagraphText(int paragraphIndex, string normalizedText)
        {
            var para = GetParagraph(paragraphIndex);
            var range = para.Range;

            string source = StripParagraphMark(range.Text);
            string target = normalizedText ?? string.Empty;
            if (string.Equals(source, target, StringComparison.Ordinal)) return;

            var edits = BuildEdits(source, target);
            if (edits.Count == 0) return;

            int paraStart = range.Start;

            // מהסוף להתחלה, כדי שהמיקומים לא יזוזו תוך כדי
            for (int e = edits.Count - 1; e >= 0; e--)
            {
                var edit = edits[e];
                var r = _doc.Range(paraStart + edit.Position, paraStart + edit.Position + 1);

                if (edit.Replacement == null) r.Delete();
                else r.Text = edit.Replacement;
            }
        }

        private struct Edit
        {
            public int Position;
            public string Replacement; // null = מחיקה
        }

        /// <summary>
        /// משווה את הטקסט המקורי למנוקה ומחזיר את רשימת העריכות.
        /// הניקוי מסיר או מחליף תווי רווח בלבד, ולכן די בהליכה אחת קדימה.
        /// </summary>
        private static List<Edit> BuildEdits(string source, string target)
        {
            var edits = new List<Edit>();
            int t = 0;

            for (int s = 0; s < source.Length; s++)
            {
                char c = source[s];

                if (t < target.Length && target[t] == c) { t++; continue; }

                if (t < target.Length && IsWhite(c) && IsWhite(target[t]))
                {
                    edits.Add(new Edit { Position = s, Replacement = target[t].ToString() });
                    t++;
                    continue;
                }

                if (IsWhite(c))
                {
                    edits.Add(new Edit { Position = s, Replacement = null });
                    continue;
                }

                // תו שאינו רווח שאינו תואם — הניקוי חרג מסמכותו. עוצרים.
                throw new InvalidOperationException(
                    "הניקוי ניסה לשנות תו שאינו רווח בפסקה. הפעולה נעצרה.");
            }

            return edits;
        }

        private static bool IsWhite(char c)
        {
            return c == ' ' || c == '\t' || c == ' ';
        }

        // ---------- מחיקה ----------

        public void DeleteParagraphs(IEnumerable<int> paragraphIndexes)
        {
            var sorted = new List<int>(paragraphIndexes);
            sorted.Sort();
            sorted.Reverse();

            int total = _doc.Paragraphs.Count;

            foreach (int i in sorted)
            {
                if (i < 0 || i >= total) continue;
                if (_doc.Paragraphs.Count <= 1) break; // וורד דורש פסקה אחת לפחות

                try { _doc.Paragraphs[i + 1].Range.Delete(); }
                catch { /* פסקה בתוך מבנה שאינו מאפשר מחיקה */ }
            }
        }

        private Word.Paragraph GetParagraph(int index)
        {
            if (index < 0 || index >= _doc.Paragraphs.Count)
                throw new ArgumentOutOfRangeException("index", "אינדקס פסקה מחוץ לטווח: " + index);
            return _doc.Paragraphs[index + 1]; // וורד מונה מ-1
        }

        // ---------- אצווה וביטול ----------

        public void BeginBatch(string undoName)
        {
            _savedScreenUpdating = _app.ScreenUpdating;
            _app.ScreenUpdating = false;

            try
            {
                _app.UndoRecord.StartCustomRecord(undoName);
                _recordOpen = true;
            }
            catch
            {
                // UndoRecord אינו זמין בגרסאות ישנות — הביטול יהיה מרובה צעדים
                _recordOpen = false;
            }
        }

        public void EndBatch()
        {
            CloseRecord();
            _app.ScreenUpdating = _savedScreenUpdating;
        }

        public void RollbackBatch()
        {
            CloseRecord();
            try
            {
                // רשומת ביטול אחת = ביטול אחד מחזיר את המסמך למצבו הקודם
                _doc.Undo(1);
            }
            catch { }
            finally
            {
                _app.ScreenUpdating = _savedScreenUpdating;
            }
        }

        private void CloseRecord()
        {
            if (!_recordOpen) return;
            try { _app.UndoRecord.EndCustomRecord(); } catch { }
            _recordOpen = false;
        }

        public void Save()
        {
            _doc.Save();
        }

        public string FullName
        {
            get { try { return _doc.FullName; } catch { return null; } }
        }

        public void Dispose() { }
    }
}
