using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using Sadar.Core.Model;

namespace Sadar.Core.Safety
{
    /// <summary>
    /// חוזה הבטיחות של המוצר.
    ///
    /// לפני כל ריצה נלקחת טביעת אצבע של המלל, ואחריה היא נלקחת שוב.
    /// אם המלל השתנה ולו בתו אחד — הריצה מתבטלת במלואה.
    ///
    /// הנרמול מתעלם ממה שהתוסף רשאי לשנות (רווחים כפולים, פסקאות ריקות,
    /// רווחים בקצוות) ורגיש לכל השאר: אותיות, ניקוד, פיסוק, גרשיים, מקפים.
    /// </summary>
    public static class TextHasher
    {
        /// <summary>נרמול פסקה יחידה לצורך השוואה.</summary>
        public static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            var sb = new StringBuilder(text.Length);
            bool lastWasSpace = true; // true בהתחלה כדי לבלוע רווחים מובילים

            foreach (char c in text)
            {
                // תווי בקרה של וורד: סוף פסקה, סוף תא, סימון שדה וכו'
                if (c == '\r' || c == '\n' || c == '\a' || c == '\v' || c == '\f' || c == '\0')
                    continue;

                // רווח, טאב, רווח קשיח, רווח באפס רוחב, רווח דק
                bool isSpace = c == ' ' || c == '	' || c == ' ' || c == '​' || c == ' ';
                if (isSpace)
                {
                    if (!lastWasSpace) { sb.Append(' '); lastWasSpace = true; }
                }
                else
                {
                    sb.Append(c);
                    lastWasSpace = false;
                }
            }

            // הנרמול חייב להיות אגרסיבי לפחות כמו המנקה, אחרת ניקוי לגיטימי
            // ייראה כשינוי מלל והריצה תתבטל בלי סיבה.
            // כאן זה מוחל תמיד, גם כשהכלל כבוי אצל המשתמש — כך כל תת-קבוצה
            // של עריכות רווח מגיעה לאותה צורה מנורמלת.
            RemoveSpaceBeforePunctuation(sb);

            // רווח עוקב בסוף
            while (sb.Length > 0 && sb[sb.Length - 1] == ' ') sb.Length--;
            return sb.ToString();
        }

        /// <summary>
        /// פיסוק שרווח לפניו נחשב שגיאת הקלדה ומותר להסירו.
        /// גרשיים אינם ברשימה: בעברית הם משמשים גם לראשי תיבות באמצע מילה.
        /// המנקה והמנרמל חייבים להשתמש באותה רשימה בדיוק.
        /// </summary>
        public static bool IsClosingPunctuation(char c)
        {
            return c == '.' || c == ',' || c == ':' || c == ';' || c == '?' || c == '!' ||
                   c == ')' || c == ']';
        }

        public static void RemoveSpaceBeforePunctuation(StringBuilder sb)
        {
            for (int i = sb.Length - 1; i > 0; i--)
            {
                if (!IsClosingPunctuation(sb[i])) continue;

                int j = i - 1;
                while (j >= 0 && sb[j] == ' ') j--;
                if (j < i - 1) sb.Remove(j + 1, i - 1 - j);
            }
        }

        /// <summary>
        /// טביעת אצבע של המסמך כולו: כל הפסקאות הלא-ריקות, מנורמלות,
        /// משורשרות בסדר הופעתן. פסקאות ריקות אינן נספרות — מחיקתן מותרת.
        /// </summary>
        public static string HashDocument(IEnumerable<ParagraphInfo> paragraphs)
        {
            var sb = new StringBuilder(1 << 16);
            foreach (var p in paragraphs)
            {
                string n = Normalize(p.Text);
                if (n.Length == 0) continue;
                sb.Append(n).Append('\n');
            }
            return Sha256(sb.ToString());
        }

        /// <summary>
        /// רשימת טביעות אצבע פר-פסקה. משמשת לאיתור מדויק של הפסקה שהשתנתה
        /// כשהאימות הכללי נכשל — כדי שההודעה למשתמש תהיה "פסקה 412" ולא "משהו השתנה".
        /// </summary>
        public static List<string> FingerprintEach(IEnumerable<ParagraphInfo> paragraphs)
        {
            var list = new List<string>();
            foreach (var p in paragraphs)
            {
                string n = Normalize(p.Text);
                list.Add(n.Length == 0 ? string.Empty : Sha256(n));
            }
            return list;
        }

        public static string Sha256(string s)
        {
            using (var sha = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(s);
                byte[] hash = sha.ComputeHash(bytes);
                var sb = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }
    }

    /// <summary>תוצאת אימות המלל.</summary>
    public sealed class VerificationResult
    {
        public bool Passed;
        public string HashBefore;
        public string HashAfter;
        public int FirstChangedParagraph = -1;
        public string Message = string.Empty;

        public static VerificationResult Ok(string hash)
        {
            return new VerificationResult
            {
                Passed = true,
                HashBefore = hash,
                HashAfter = hash,
                Message = "המלל אומת — לא נמצא שינוי."
            };
        }
    }

    /// <summary>מבצע את ההשוואה ומאתר את נקודת השינוי הראשונה.</summary>
    public static class TextVerifier
    {
        public static VerificationResult Verify(
            List<ParagraphInfo> before,
            List<ParagraphInfo> after)
        {
            string hb = TextHasher.HashDocument(before);
            string ha = TextHasher.HashDocument(after);

            var result = new VerificationResult { HashBefore = hb, HashAfter = ha };

            if (string.Equals(hb, ha, StringComparison.Ordinal))
            {
                result.Passed = true;
                result.Message = "המלל אומת — לא נמצא שינוי.";
                return result;
            }

            result.Passed = false;

            // איתור הפסקה הראשונה שנשתנתה, תוך התעלמות מפסקאות ריקות שנמחקו
            var fb = NonEmptyNormalized(before);
            var fa = NonEmptyNormalized(after);
            int n = Math.Min(fb.Count, fa.Count);
            for (int i = 0; i < n; i++)
            {
                if (!string.Equals(fb[i].Value, fa[i].Value, StringComparison.Ordinal))
                {
                    result.FirstChangedParagraph = fb[i].Key;
                    result.Message = string.Format(
                        "המלל השתנה בפסקה {0}. הפעולה בוטלה במלואה.", fb[i].Key + 1);
                    return result;
                }
            }

            if (fb.Count != fa.Count)
            {
                result.FirstChangedParagraph = n > 0 ? fb[Math.Min(n, fb.Count - 1)].Key : 0;
                result.Message = string.Format(
                    "מספר הפסקאות עם תוכן השתנה ({0} במקום {1}). הפעולה בוטלה במלואה.",
                    fa.Count, fb.Count);
                return result;
            }

            result.Message = "נמצא שינוי במלל. הפעולה בוטלה במלואה.";
            return result;
        }

        private static List<KeyValuePair<int, string>> NonEmptyNormalized(List<ParagraphInfo> ps)
        {
            var list = new List<KeyValuePair<int, string>>(ps.Count);
            foreach (var p in ps)
            {
                string n = TextHasher.Normalize(p.Text);
                if (n.Length > 0) list.Add(new KeyValuePair<int, string>(p.Index, n));
            }
            return list;
        }
    }
}
