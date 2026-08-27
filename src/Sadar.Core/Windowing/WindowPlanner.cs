using System;
using System.Collections.Generic;
using Sadar.Core.Model;
using Sadar.Core.Rules;

namespace Sadar.Core.Windowing
{
    public sealed class Window
    {
        public int Start;
        public int Count;
        public int OverlapStart; // אינדקס הפסקה הראשונה שגם החלון הקודם ראה

        public int End { get { return Start + Count; } }
    }

    /// <summary>
    /// חלוקת מסמך ארוך לחלונות חופפים.
    ///
    /// זה מה שפותר את המגבלה שכולם בשרשור נתקלו בה — ספר של מאות עמודים
    /// שלא נכנס בבת אחת. המשתמש לא מפצל שום דבר; זה קורה מתחת למכסה המנוע.
    /// </summary>
    public static class WindowPlanner
    {
        public static List<Window> Plan(int paragraphCount, RulesFile rules)
        {
            var windows = new List<Window>();
            if (paragraphCount <= 0) return windows;

            int size = Math.Max(50, rules.WindowSize);
            int overlap = Math.Max(0, Math.Min(rules.WindowOverlap, size / 2));

            if (paragraphCount <= size)
            {
                windows.Add(new Window { Start = 0, Count = paragraphCount, OverlapStart = -1 });
                return windows;
            }

            int pos = 0;
            while (pos < paragraphCount)
            {
                int count = Math.Min(size, paragraphCount - pos);
                windows.Add(new Window
                {
                    Start = pos,
                    Count = count,
                    OverlapStart = pos == 0 ? -1 : pos
                });

                if (pos + count >= paragraphCount) break;
                pos += (size - overlap);
            }

            return windows;
        }
    }

    /// <summary>
    /// מיזוג ההחלטות מכל החלונות.
    ///
    /// באזור החפיפה שני חלונות מחווים דעה על אותה פסקה. אם הם חלוקים —
    /// הפסקה עוברת לבדיקת אדם. לא בוחרים שרירותית ולא מכריעים ברוב.
    /// </summary>
    public static class DecisionMerger
    {
        public static DecisionSet Merge(IEnumerable<DecisionSet> windowResults)
        {
            var merged = new DecisionSet();
            var byIndex = new Dictionary<int, StyleDecision>();
            var conflicts = new HashSet<int>();
            var deletes = new HashSet<int>();

            foreach (var set in windowResults)
            {
                if (set == null) continue;

                merged.Warnings.AddRange(set.Warnings);

                foreach (int d in set.Deletes) deletes.Add(d);

                foreach (var dec in set.Decisions)
                {
                    StyleDecision existing;
                    if (!byIndex.TryGetValue(dec.Index, out existing))
                    {
                        byIndex[dec.Index] = dec;
                        continue;
                    }

                    if (string.Equals(existing.StyleName, dec.StyleName, StringComparison.Ordinal))
                    {
                        // הסכמה בין חלונות מחזקת — אם אחד היה בטוח, נשארים בטוחים
                        if (!dec.IsUncertain) existing.IsUncertain = false;
                        continue;
                    }

                    conflicts.Add(dec.Index);
                    existing.IsUncertain = true;
                    existing.Reason = string.Format(
                        "אי-הסכמה בין חלונות: \"{0}\" מול \"{1}\". נדרשת הכרעה.",
                        existing.StyleName, dec.StyleName);
                }
            }

            var indexes = new List<int>(byIndex.Keys);
            indexes.Sort();
            foreach (int i in indexes) merged.Decisions.Add(byIndex[i]);

            var delList = new List<int>(deletes);
            delList.Sort();
            merged.Deletes.AddRange(delList);

            if (conflicts.Count > 0)
                merged.Warnings.Add(string.Format(
                    "{0} פסקאות באזורי החפיפה קיבלו הכרעות שונות וסומנו לבדיקה.", conflicts.Count));

            return merged;
        }
    }
}
