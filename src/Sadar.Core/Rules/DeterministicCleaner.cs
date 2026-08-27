using System;
using System.Collections.Generic;
using System.Text;
using Sadar.Core.Model;

namespace Sadar.Core.Rules
{
    /// <summary>
    /// ניקוי לפי חוקים יבשים — רץ מקומית לגמרי, בלי שום קריאה למודל.
    ///
    /// זה החלק שמאקרו כבר יודע לעשות, ולכן אין סיבה לשלם עליו טוקנים.
    /// המשתמש יכול להריץ רק אותו, בלי חיבור לאינטרנט בכלל.
    /// </summary>
    public static class DeterministicCleaner
    {
        public sealed class CleanPlan
        {
            /// <summary>אינדקס פסקה -> הטקסט המנורמל שיש לכתוב לה.</summary>
            public readonly Dictionary<int, string> Rewrites = new Dictionary<int, string>();

            /// <summary>אינדקסים של פסקאות ריקות שיש למחוק.</summary>
            public readonly List<int> Deletes = new List<int>();

            public int TotalChanges { get { return Rewrites.Count + Deletes.Count; } }
        }

        public static CleanPlan Plan(IList<ParagraphInfo> paragraphs, RulesFile rules)
        {
            var plan = new CleanPlan();
            if (paragraphs == null || paragraphs.Count == 0) return plan;

            // ----- ניקוי טקסט בתוך פסקאות -----
            foreach (var p in paragraphs)
            {
                if (rules.IsProtected(p)) continue;
                if (p.IsEmpty) continue;

                string cleaned = CleanText(p.Text, rules);
                if (!string.Equals(cleaned, p.Text, StringComparison.Ordinal))
                    plan.Rewrites[p.Index] = cleaned;
            }

            if (!rules.RemoveEmptyParagraphs) return plan;

            // ----- פסקאות ריקות עודפות ברצף -----
            var alreadyMarked = new HashSet<int>();

            int run = 0;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].IsEmpty)
                {
                    run++;
                    if (run > rules.MaxConsecutiveEmpty)
                    {
                        plan.Deletes.Add(paragraphs[i].Index);
                        alreadyMarked.Add(paragraphs[i].Index);
                    }
                }
                else
                {
                    run = 0;
                }
            }

            // ----- פסקאות ריקות בסוף המסמך -----
            if (rules.TrimTrailingEmpty)
            {
                for (int i = paragraphs.Count - 1; i >= 0; i--)
                {
                    if (!paragraphs[i].IsEmpty) break;
                    // חיפוש ברשימה כאן היה ריבועי במסמכים גדולים
                    if (alreadyMarked.Add(paragraphs[i].Index))
                        plan.Deletes.Add(paragraphs[i].Index);
                }
            }

            // וורד דורש שתישאר לפחות פסקה אחת במסמך
            if (plan.Deletes.Count >= paragraphs.Count && plan.Deletes.Count > 0)
                plan.Deletes.RemoveAt(plan.Deletes.Count - 1);

            return plan;
        }

        /// <summary>
        /// ניקוי רווחים בלבד. אינו מוסיף, מסיר או משנה תו שאינו רווח —
        /// זו הסיבה שאימות ה-hash עובר אחרי הניקוי.
        /// </summary>
        public static string CleanText(string text, RulesFile rules)
        {
            if (string.IsNullOrEmpty(text)) return text;

            var sb = new StringBuilder(text.Length);
            bool lastWasSpace = false;

            foreach (char c in text)
            {
                char ch = c;

                if (rules.ConvertTabsToSpaces && ch == '\t')
                    ch = ' ';

                bool isSpace = ch == ' ' || ch == ' ';

                if (isSpace && rules.CollapseDoubleSpaces)
                {
                    if (lastWasSpace) continue;
                    lastWasSpace = true;
                    sb.Append(' ');
                    continue;
                }

                lastWasSpace = isSpace;
                sb.Append(ch);
            }

            if (rules.RemoveSpaceBeforePunctuation)
                Safety.TextHasher.RemoveSpaceBeforePunctuation(sb);

            if (rules.TrimParagraphEdges)
            {
                while (sb.Length > 0 && IsTrimmable(sb[0])) sb.Remove(0, 1);
                while (sb.Length > 0 && IsTrimmable(sb[sb.Length - 1])) sb.Length--;
            }

            return sb.ToString();
        }

        private static bool IsTrimmable(char c)
        {
            return c == ' ' || c == '\t' || c == ' ';
        }

    }
}
