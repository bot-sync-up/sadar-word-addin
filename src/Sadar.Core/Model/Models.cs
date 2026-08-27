using System;
using System.Collections.Generic;

namespace Sadar.Core.Model
{
    /// <summary>יישור פסקה, מנורמל לערכים שהמודל מקבל.</summary>
    public enum Align { Right, Left, Center, Justify, Unknown }

    /// <summary>
    /// תיאור מבני של פסקה אחת. זהו כל מה שיוצא מהמסמך אל המודל —
    /// ובכוונה תחילה הוא מכיל תחילית קצרה ולא את גוף הטקסט.
    /// </summary>
    public sealed class ParagraphInfo
    {
        public int Index;
        public string Text;          // הטקסט המלא — נשאר מקומי, לא נשלח לעולם
        public string StyleName;
        public bool Bold;
        public double FontSize;
        public Align Alignment = Align.Unknown;
        public int OutlineLevel = 9; // 9 = גוף טקסט בוורד
        public bool IsListItem;

        public bool IsEmpty
        {
            get { return string.IsNullOrEmpty(Text) || Text.Trim().Length == 0; }
        }

        public int Length
        {
            get { return Text == null ? 0 : Text.Length; }
        }

        /// <summary>התחילית שנשלחת למודל. אורך קבוע ומוגבל.</summary>
        public string Prefix(int maxChars)
        {
            if (string.IsNullOrEmpty(Text)) return string.Empty;
            string t = Text.Trim();
            if (t.Length <= maxChars) return t;
            return t.Substring(0, maxChars);
        }
    }

    /// <summary>החלטת סגנון יחידה שהתקבלה מהמודל.</summary>
    public sealed class StyleDecision
    {
        public int Index;
        public string StyleName;
        public bool IsUncertain;
        public string Reason;

        public override string ToString()
        {
            return string.Format("[{0}] -> {1}{2}", Index, StyleName, IsUncertain ? " (?)" : "");
        }
    }

    /// <summary>תוצאת חלון אחד או של המסמך כולו.</summary>
    public sealed class DecisionSet
    {
        public readonly List<StyleDecision> Decisions = new List<StyleDecision>();
        public readonly List<int> Deletes = new List<int>();
        public readonly List<string> Warnings = new List<string>();

        public int Count { get { return Decisions.Count; } }

        public void Merge(DecisionSet other)
        {
            if (other == null) return;
            Decisions.AddRange(other.Decisions);
            Deletes.AddRange(other.Deletes);
            Warnings.AddRange(other.Warnings);
        }
    }

    /// <summary>מה שהתוסף עשה בפועל, לדוח הסיום וליומן.</summary>
    public sealed class RunReport
    {
        public int ParagraphsScanned;
        public int StylesApplied;
        public int ParagraphsDeleted;
        public int WhitespaceFixes;
        public int UncertainCount;
        public int WindowsProcessed;
        public bool TextVerified;
        public string BackupPath;
        public TimeSpan Duration;
        public readonly List<string> Warnings = new List<string>();
        public readonly Dictionary<string, int> StyleCounts = new Dictionary<string, int>();

        public void CountStyle(string style)
        {
            if (string.IsNullOrEmpty(style)) return;
            int n;
            StyleCounts.TryGetValue(style, out n);
            StyleCounts[style] = n + 1;
        }
    }
}
