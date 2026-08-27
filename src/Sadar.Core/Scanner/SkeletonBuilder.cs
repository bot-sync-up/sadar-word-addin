using System;
using System.Collections.Generic;
using System.Text;
using Sadar.Core.Model;
using Sadar.Core.Rules;

namespace Sadar.Core.Scanner
{
    /// <summary>
    /// בונה את השלד — הייצוג היחיד של המסמך שיוצא החוצה.
    ///
    /// כלל ברזל: גוף הטקסט לעולם אינו נכלל. רק תחילית קצרה ומאפייני עיצוב.
    /// במצב מסונן גם התחילית מוחלפת בתבנית מבנית.
    /// </summary>
    public static class SkeletonBuilder
    {
        /// <summary>
        /// ערכי ברירת המחדל של המסמך. כל פסקה שתואמת אותם משמיטה את השדה,
        /// ובמסמך אמיתי רוב הפסקאות תואמות — זה מה שמכווץ את השלד.
        /// </summary>
        public sealed class Defaults
        {
            public string Style = "רגיל";
            public double FontSize;
            public Align Alignment = Align.Justify;

            public static Defaults Detect(IList<ParagraphInfo> paragraphs)
            {
                var styles = new Dictionary<string, int>();
                var sizes = new Dictionary<double, int>();
                var aligns = new Dictionary<Align, int>();

                foreach (var p in paragraphs)
                {
                    if (p.IsEmpty) continue;
                    Bump(styles, p.StyleName ?? "רגיל");
                    if (p.FontSize > 0) Bump(sizes, p.FontSize);
                    Bump(aligns, p.Alignment);
                }

                return new Defaults
                {
                    Style = Top(styles, "רגיל"),
                    FontSize = Top(sizes, 0.0),
                    Alignment = Top(aligns, Align.Justify)
                };
            }

            private static void Bump<T>(Dictionary<T, int> d, T key)
            {
                int n;
                d.TryGetValue(key, out n);
                d[key] = n + 1;
            }

            private static T Top<T>(Dictionary<T, int> d, T fallback)
            {
                int best = -1;
                T result = fallback;
                foreach (var kv in d)
                    if (kv.Value > best) { best = kv.Value; result = kv.Key; }
                return result;
            }
        }

        /// <summary>
        /// מעל אורך זה פסקה היא גוף טקסט מובהק ואין שום ערך בתחילית שלה.
        /// נשלח רק האורך. חוסך נפח, ובעיקר — הטקסט הארוך לא יוצא מהמחשב כלל.
        /// </summary>
        public const int ObviousBodyLength = 180;

        public static string Build(
            IList<ParagraphInfo> paragraphs,
            int from,
            int count,
            RulesFile rules,
            string runningContext)
        {
            var def = Defaults.Detect(paragraphs);

            var w = new Json.Json.Writer();
            w.StartObject();

            w.Prop("structureHint", rules.StructureHint);

            w.Name("allowedStyles").StartArray();
            foreach (var s in rules.AllowedStyles) w.Value(s);
            w.EndArray();

            if (!string.IsNullOrEmpty(runningContext))
                w.Prop("contextSoFar", runningContext);

            if (rules.PrivacyMode) w.Prop("privacyMode", true);

            // הסבר הקיצורים — שדה שחסר בפסקה משמעו "כערך ברירת המחדל"
            w.Name("defaults").StartObject()
                .Prop("s", def.Style)
                .Prop("al", AlignCode(def.Alignment));
            if (def.FontSize > 0) w.Prop("sz", def.FontSize);
            w.EndObject();

            w.Prop("note",
                "שדה שחסר בפסקה = כערך ב-defaults. פסקה עם len בלבד היא גוף טקסט ארוך " +
                "שתחיליתו לא נשלחה. b=מודגש, len=אורך מלא בתווים. " +
                "emptyIndexes הם אינדקסים של פסקאות ריקות ואינם מופיעים ברשימת paragraphs.");

            int end = Math.Min(from + count, paragraphs.Count);

            // פסקאות ריקות נשלחות כרשימת אינדקסים ולא כאובייקטים.
            // במסמך אמיתי הן כמחצית מהפסקאות, וזה ההבדל בין שלד כבד לקל.
            w.Name("emptyIndexes").StartArray();
            for (int i = from; i < end; i++)
                if (paragraphs[i].IsEmpty) w.Value(paragraphs[i].Index);
            w.EndArray();

            w.Name("paragraphs").StartArray();

            for (int i = from; i < end; i++)
            {
                var p = paragraphs[i];
                if (p.IsEmpty) continue;

                w.StartObject();
                w.Prop("i", p.Index);

                bool obviousBody = p.Length > ObviousBodyLength && !p.Bold;

                if (!obviousBody)
                {
                    if (rules.PrivacyMode)
                        w.Prop("pattern", Fingerprint(p.Text));
                    else
                        w.Prop("t", p.Prefix(rules.PrefixChars));
                }

                if (!string.Equals(p.StyleName ?? "רגיל", def.Style, StringComparison.Ordinal))
                    w.Prop("s", p.StyleName ?? "רגיל");

                if (p.Bold) w.Prop("b", true);

                if (p.FontSize > 0 && Math.Abs(p.FontSize - def.FontSize) > 0.01)
                    w.Prop("sz", p.FontSize);

                if (p.Alignment != def.Alignment)
                    w.Prop("al", AlignCode(p.Alignment));

                w.Prop("len", p.Length);

                if (p.IsListItem) w.Prop("list", true);
                if (p.OutlineLevel < 9) w.Prop("lvl", p.OutlineLevel);

                w.EndObject();
            }

            w.EndArray();
            w.EndObject();
            return w.ToString();
        }

        private static string AlignCode(Align a)
        {
            switch (a)
            {
                case Align.Right: return "right";
                case Align.Left: return "left";
                case Align.Center: return "center";
                case Align.Justify: return "justify";
                default: return "?";
            }
        }

        /// <summary>
        /// תבנית מבנית למצב מסונן: מספר מילים, סוג התו הפותח והסוגר,
        /// ונוכחות סימני פיסוק — בלי אף אות מהטקסט עצמו.
        /// לדוגמה: "סימן א" -> "W2|heb|heb|no-punct"
        /// </summary>
        public static string Fingerprint(string text)
        {
            if (string.IsNullOrEmpty(text)) return "empty";
            string t = text.Trim();
            if (t.Length == 0) return "empty";

            int words = 1;
            for (int i = 0; i < t.Length; i++)
                if (t[i] == ' ') words++;

            var sb = new StringBuilder(32);
            sb.Append('W').Append(words);
            sb.Append('|').Append(CharClass(t[0]));
            sb.Append('|').Append(CharClass(t[t.Length - 1]));

            bool hasPunct = false, hasQuote = false, hasDigit = false;
            foreach (char c in t)
            {
                if (c == '.' || c == ',' || c == ':' || c == ';' || c == '?' || c == '!') hasPunct = true;
                else if (c == '"' || c == '\'' || c == '״' || c == '׳') hasQuote = true;
                else if (c >= '0' && c <= '9') hasDigit = true;
            }
            if (hasPunct) sb.Append("|punct");
            if (hasQuote) sb.Append("|quote");
            if (hasDigit) sb.Append("|digit");

            return sb.ToString();
        }

        private static string CharClass(char c)
        {
            if (c >= 'א' && c <= 'ת') return "heb";
            if (c >= '֐' && c <= '״') return "hebmark";
            if (char.IsDigit(c)) return "num";
            if (char.IsLetter(c)) return "lat";
            if (char.IsPunctuation(c)) return "punct";
            return "other";
        }

        /// <summary>
        /// תקציר מבני מצטבר שמועבר לחלון הבא, כדי שהחלטות יישארו עקביות
        /// לאורך ספר שלם ולא ישתנו בין פרק לפרק.
        /// </summary>
        public static string BuildRunningContext(
            IList<ParagraphInfo> paragraphs,
            DecisionSet decisionsSoFar,
            RulesFile rules)
        {
            if (decisionsSoFar == null || decisionsSoFar.Count == 0) return null;

            var counts = new Dictionary<string, int>();
            var examples = new Dictionary<string, string>();

            foreach (var d in decisionsSoFar.Decisions)
            {
                if (string.Equals(d.StyleName, rules.BodyStyle, StringComparison.Ordinal)) continue;

                int n;
                counts.TryGetValue(d.StyleName, out n);
                counts[d.StyleName] = n + 1;

                if (!examples.ContainsKey(d.StyleName))
                {
                    var p = FindByIndex(paragraphs, d.Index);
                    if (p != null && !p.IsEmpty)
                        examples[d.StyleName] = p.Prefix(40);
                }
            }

            if (counts.Count == 0) return null;

            var sb = new StringBuilder();
            sb.Append("עד כה זוהו: ");
            bool first = true;
            foreach (var kv in counts)
            {
                if (!first) sb.Append("; ");
                first = false;
                sb.Append(kv.Value).Append(" x ").Append(kv.Key);
                string ex;
                if (examples.TryGetValue(kv.Key, out ex))
                    sb.Append(" (למשל: \"").Append(ex).Append("\")");
            }
            sb.Append(". שמור על אותה הבחנה בהמשך.");
            return sb.ToString();
        }

        private static ParagraphInfo FindByIndex(IList<ParagraphInfo> ps, int index)
        {
            // הפסקאות ממוינות לפי אינדקס, אז חיפוש בינארי
            int lo = 0, hi = ps.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                int v = ps[mid].Index;
                if (v == index) return ps[mid];
                if (v < index) lo = mid + 1; else hi = mid - 1;
            }
            return null;
        }
    }
}
