using System;
using System.Collections;
using System.Collections.Generic;
using Sadar.Core.Model;
using Sadar.Core.Rules;

namespace Sadar.Core.Engine
{
    /// <summary>
    /// ולידציה קשיחה של התשובה מהמודל.
    ///
    /// כל דבר שאינו עומד בסכמה נפסל ונרשם ביומן, ולא מוחל על המסמך.
    /// המודל הוא מקור המלצה — לא מקור סמכות.
    /// </summary>
    public static class DecisionParser
    {
        public static DecisionSet Parse(
            string json,
            IDictionary<int, ParagraphInfo> allowedParagraphs,
            RulesFile rules)
        {
            var set = new DecisionSet();
            if (string.IsNullOrEmpty(json))
            {
                set.Warnings.Add("התקבלה תשובה ריקה מהמנוע.");
                return set;
            }

            IDictionary root;
            try
            {
                root = Json.Json.ParseObject(json);
            }
            catch (Exception ex)
            {
                set.Warnings.Add("התשובה מהמנוע אינה JSON תקין: " + ex.Message);
                return set;
            }

            var allowedStyles = new HashSet<string>(rules.AllowedStyles, StringComparer.Ordinal);
            var seen = new HashSet<int>();

            // ----- decisions -----
            foreach (var item in Json.Json.GetArray(root, "decisions"))
            {
                var d = item as IDictionary;
                if (d == null) continue;

                int i = Json.Json.GetInt(d, "i", -1);
                string style = Json.Json.GetString(d, "style", null);

                if (!Validate(i, style, allowedParagraphs, allowedStyles, rules, set, seen)) continue;

                set.Decisions.Add(new StyleDecision { Index = i, StyleName = style });
                seen.Add(i);
            }

            // ----- uncertain -----
            foreach (var item in Json.Json.GetArray(root, "uncertain"))
            {
                var d = item as IDictionary;
                if (d == null) continue;

                int i = Json.Json.GetInt(d, "i", -1);
                string style = Json.Json.GetString(d, "style", null);
                string why = Json.Json.GetString(d, "why", "לא צוינה סיבה");

                // פסקה מסופקת בלי הצעת סגנון — מסמנים אותה לבדיקה בלי החלטה
                if (string.IsNullOrEmpty(style)) style = rules.BodyStyle;

                if (!Validate(i, style, allowedParagraphs, allowedStyles, rules, set, seen)) continue;

                set.Decisions.Add(new StyleDecision
                {
                    Index = i,
                    StyleName = style,
                    IsUncertain = true,
                    Reason = why
                });
                seen.Add(i);
            }

            // ----- deletes -----
            foreach (var item in Json.Json.GetArray(root, "deletes"))
            {
                int i;
                try { i = Convert.ToInt32(item, System.Globalization.CultureInfo.InvariantCulture); }
                catch { continue; }

                ParagraphInfo p;
                if (!allowedParagraphs.TryGetValue(i, out p))
                {
                    set.Warnings.Add(string.Format("נפסלה מחיקה של פסקה {0}: אינה בטווח החלון.", i));
                    continue;
                }

                // ההגנה המרכזית: מחיקה מותרת אך ורק על פסקה ריקה
                if (!p.IsEmpty)
                {
                    set.Warnings.Add(string.Format(
                        "נפסלה מחיקה של פסקה {0}: הפסקה מכילה טקסט. מחיקה מותרת רק לפסקאות ריקות.", i + 1));
                    continue;
                }

                if (rules.IsProtected(p))
                {
                    set.Warnings.Add(string.Format("נפסלה מחיקה של פסקה {0}: מוגנת לפי הכללים.", i + 1));
                    continue;
                }

                set.Deletes.Add(i);
            }

            return set;
        }

        private static bool Validate(
            int index,
            string style,
            IDictionary<int, ParagraphInfo> allowed,
            HashSet<string> allowedStyles,
            RulesFile rules,
            DecisionSet set,
            HashSet<int> seen)
        {
            if (index < 0)
            {
                set.Warnings.Add("נפסלה החלטה ללא אינדקס פסקה תקין.");
                return false;
            }

            if (seen.Contains(index)) return false; // כפילות — הראשונה קובעת

            ParagraphInfo p;
            if (!allowed.TryGetValue(index, out p))
            {
                set.Warnings.Add(string.Format("נפסלה החלטה על פסקה {0}: אינה בטווח החלון.", index));
                return false;
            }

            if (string.IsNullOrEmpty(style) || !allowedStyles.Contains(style))
            {
                set.Warnings.Add(string.Format(
                    "נפסלה החלטה על פסקה {0}: הסגנון \"{1}\" אינו ברשימת הסגנונות המותרים.",
                    index + 1, style ?? "(ריק)"));
                return false;
            }

            if (p.IsEmpty)
            {
                set.Warnings.Add(string.Format("נפסלה החלטה על פסקה {0}: הפסקה ריקה.", index + 1));
                return false;
            }

            if (rules.IsProtected(p))
            {
                set.Warnings.Add(string.Format("דולגה פסקה {0}: מוגנת לפי הכללים.", index + 1));
                return false;
            }

            return true;
        }
    }
}
