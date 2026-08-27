using System;
using System.Collections.Generic;
using Sadar.Core.Model;
using Sadar.Core.Rules;

namespace Sadar.Core.Engine
{
    /// <summary>
    /// מנוע ההחלטות. מקבל שלד ומחזיר החלטות סגנון.
    ///
    /// הממשק מכוון להיות צר בכוונה: נכנס טקסט אחד, יוצא JSON אחד.
    /// אין לו גישה למסמך, לדיסק או לוורד.
    /// </summary>
    public interface IDecisionEngine
    {
        string Name { get; }

        /// <summary>בדיקה שהמנוע זמין ומאומת. מחזיר הודעת שגיאה בעברית, או null אם תקין.</summary>
        string CheckAvailability();

        /// <summary>שולח שלד ומקבל JSON גולמי. זורק חריגה על כשל.</summary>
        string Decide(string skeletonJson, RulesFile rules);
    }

    /// <summary>שגיאה שמקורה במנוע, עם הודעה מוכנה למשתמש בעברית.</summary>
    public sealed class EngineException : Exception
    {
        public readonly bool LooksLikeFilterBlock;

        public EngineException(string message, bool filterBlock = false, Exception inner = null)
            : base(message, inner)
        {
            LooksLikeFilterBlock = filterBlock;
        }
    }

    /// <summary>
    /// מנוע מדומה לפיתוח ולבדיקות — מחליט לפי היוריסטיקה מקומית,
    /// בלי רשת ובלי עלות. מאפשר לבדוק את כל הצינור מקצה לקצה.
    /// </summary>
    public sealed class MockEngine : IDecisionEngine
    {
        private readonly IList<ParagraphInfo> _paragraphs;

        public MockEngine(IList<ParagraphInfo> paragraphs)
        {
            _paragraphs = paragraphs;
        }

        public string Name { get { return "מנוע מדומה (בדיקות)"; } }

        public string CheckAvailability() { return null; }

        public string Decide(string skeletonJson, RulesFile rules)
        {
            var d = Json.Json.ParseObject(skeletonJson);

            // שדה חסר בפסקה משמעו "כערך ברירת המחדל" — המנוע המדומה
            // חייב לקרוא את השלד באותם כללים שהמודל האמיתי מקבל
            var defaults = d.ContainsKey("defaults")
                ? d["defaults"] as System.Collections.IDictionary
                : null;
            string defaultAlign = defaults == null ? "?" : Json.Json.GetString(defaults, "al", "?");

            var w = new Json.Json.Writer();
            w.StartObject();
            w.Name("decisions").StartArray();

            var deletes = new List<int>();
            foreach (var e in Json.Json.GetArray(d, "emptyIndexes"))
            {
                try { deletes.Add(Convert.ToInt32(e, System.Globalization.CultureInfo.InvariantCulture)); }
                catch { }
            }

            string h1 = Pick(rules, 0), h2 = Pick(rules, 1);

            foreach (var item in Json.Json.GetArray(d, "paragraphs"))
            {
                var p = item as System.Collections.IDictionary;
                if (p == null) continue;

                int i = Json.Json.GetInt(p, "i", -1);
                if (i < 0) continue;

                string t = Json.Json.GetString(p, "t", "");
                bool bold = Json.Json.GetBool(p, "b");
                int len = Json.Json.GetInt(p, "len", 0);
                string align = Json.Json.GetString(p, "al", defaultAlign);

                string style;
                if (bold && len <= 12 && t.StartsWith("סימן", StringComparison.Ordinal))
                    style = h1;
                else if (bold && len <= 60 && align == "center")
                    style = h2;
                else
                    style = rules.BodyStyle;

                w.StartObject().Prop("i", i).Prop("style", style).EndObject();
            }

            w.EndArray();
            w.Name("deletes").StartArray();
            foreach (int i in deletes) w.Value(i);
            w.EndArray();
            w.Name("uncertain").StartArray().EndArray();
            w.EndObject();
            return w.ToString();
        }

        private static string Pick(RulesFile r, int i)
        {
            if (r.AllowedStyles.Count > i) return r.AllowedStyles[i];
            return r.BodyStyle;
        }
    }
}
