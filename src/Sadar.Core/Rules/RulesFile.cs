using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Sadar.Core.Json;

namespace Sadar.Core.Rules
{
    /// <summary>
    /// קובץ הכללים — הלב של העבודה החוזרת.
    ///
    /// מעמד שעובד על אותה סדרה שוב ושוב מגדיר פעם אחת איך הספר בנוי,
    /// ומכאן כל ספר בסדרה מסודר באותם כללים בדיוק. זה בדיוק מה שאנשים
    /// בשרשור בנו לעצמם ידנית.
    /// </summary>
    public sealed class RulesFile
    {
        public string Name = "כללי ברירת מחדל";
        public string Description = "";

        /// <summary>ההגדרה המבנית במילים של המשתמש. זה מה שהמודל מקבל.</summary>
        public string StructureHint =
            "הספר מחולק לסימנים. שם הסימן הוא כותרת ראשית, והכותרת התיאורית שאחריו היא כותרת משנה.";

        /// <summary>הסגנונות שמותר למודל לבחור מהם. כל דבר אחר נפסל בוולידציה.</summary>
        public List<string> AllowedStyles = new List<string>
        {
            "כותרת 1", "כותרת 2", "כותרת 3", "רגיל"
        };

        /// <summary>שם הסגנון שמשמעו "אל תיגע" — פסקאות שנשארות כפי שהן.</summary>
        public string BodyStyle = "רגיל";

        // ----- ניקוי דטרמיניסטי -----
        public bool CollapseDoubleSpaces = true;
        public bool TrimParagraphEdges = true;
        public bool RemoveSpaceBeforePunctuation = true;
        public bool ConvertTabsToSpaces = true;
        public bool RemoveEmptyParagraphs = true;

        /// <summary>כמה פסקאות ריקות מותר להשאיר ברצף. 0 = למחוק את כולן.</summary>
        public int MaxConsecutiveEmpty = 0;

        /// <summary>מחיקת פסקאות ריקות בסוף המסמך.</summary>
        public bool TrimTrailingEmpty = true;

        // ----- החרגות -----
        /// <summary>פסקאות שהתחילית שלהן מתחילה באחד מאלה לא ייגעו כלל.</summary>
        public List<string> NeverTouchPrefixes = new List<string>();

        /// <summary>סגנונות שאם פסקה כבר נמצאת בהם — משאירים אותה.</summary>
        public List<string> PreserveExistingStyles = new List<string>();

        // ----- שליחה למודל -----
        /// <summary>כמה תווים מכל פסקה נשלחים. ככל שפחות — פחות חשיפה, פחות דיוק.</summary>
        public int PrefixChars = 60;

        /// <summary>מצב מסונן: שולח תבנית מבנית במקום טקסט.</summary>
        public bool PrivacyMode = false;

        public int WindowSize = 400;
        public int WindowOverlap = 30;

        // ---------- שמירה וטעינה ----------

        public string ToJson()
        {
            var w = new Json.Json.Writer();
            w.StartObject();
            w.Prop("name", Name);
            w.Prop("description", Description);
            w.Prop("structureHint", StructureHint);

            w.Name("allowedStyles").StartArray();
            foreach (var s in AllowedStyles) w.Value(s);
            w.EndArray();

            w.Prop("bodyStyle", BodyStyle);
            w.Prop("collapseDoubleSpaces", CollapseDoubleSpaces);
            w.Prop("trimParagraphEdges", TrimParagraphEdges);
            w.Prop("removeSpaceBeforePunctuation", RemoveSpaceBeforePunctuation);
            w.Prop("convertTabsToSpaces", ConvertTabsToSpaces);
            w.Prop("removeEmptyParagraphs", RemoveEmptyParagraphs);
            w.Prop("maxConsecutiveEmpty", MaxConsecutiveEmpty);
            w.Prop("trimTrailingEmpty", TrimTrailingEmpty);

            w.Name("neverTouchPrefixes").StartArray();
            foreach (var s in NeverTouchPrefixes) w.Value(s);
            w.EndArray();

            w.Name("preserveExistingStyles").StartArray();
            foreach (var s in PreserveExistingStyles) w.Value(s);
            w.EndArray();

            w.Prop("prefixChars", PrefixChars);
            w.Prop("privacyMode", PrivacyMode);
            w.Prop("windowSize", WindowSize);
            w.Prop("windowOverlap", WindowOverlap);
            w.EndObject();
            return w.ToString();
        }

        public static RulesFile FromJson(string json)
        {
            var d = Json.Json.ParseObject(json);
            var r = new RulesFile();

            r.Name = Json.Json.GetString(d, "name", r.Name);
            r.Description = Json.Json.GetString(d, "description", r.Description);
            r.StructureHint = Json.Json.GetString(d, "structureHint", r.StructureHint);

            var styles = ReadStringList(d, "allowedStyles");
            if (styles.Count > 0) r.AllowedStyles = styles;

            r.BodyStyle = Json.Json.GetString(d, "bodyStyle", r.BodyStyle);
            r.CollapseDoubleSpaces = Json.Json.GetBool(d, "collapseDoubleSpaces", r.CollapseDoubleSpaces);
            r.TrimParagraphEdges = Json.Json.GetBool(d, "trimParagraphEdges", r.TrimParagraphEdges);
            r.RemoveSpaceBeforePunctuation = Json.Json.GetBool(d, "removeSpaceBeforePunctuation", r.RemoveSpaceBeforePunctuation);
            r.ConvertTabsToSpaces = Json.Json.GetBool(d, "convertTabsToSpaces", r.ConvertTabsToSpaces);
            r.RemoveEmptyParagraphs = Json.Json.GetBool(d, "removeEmptyParagraphs", r.RemoveEmptyParagraphs);
            r.MaxConsecutiveEmpty = Json.Json.GetInt(d, "maxConsecutiveEmpty", r.MaxConsecutiveEmpty);
            r.TrimTrailingEmpty = Json.Json.GetBool(d, "trimTrailingEmpty", r.TrimTrailingEmpty);

            r.NeverTouchPrefixes = ReadStringList(d, "neverTouchPrefixes");
            r.PreserveExistingStyles = ReadStringList(d, "preserveExistingStyles");

            r.PrefixChars = Clamp(Json.Json.GetInt(d, "prefixChars", r.PrefixChars), 10, 200);
            r.PrivacyMode = Json.Json.GetBool(d, "privacyMode", r.PrivacyMode);
            r.WindowSize = Clamp(Json.Json.GetInt(d, "windowSize", r.WindowSize), 50, 2000);
            r.WindowOverlap = Clamp(Json.Json.GetInt(d, "windowOverlap", r.WindowOverlap), 0, r.WindowSize / 2);

            return r;
        }

        private static List<string> ReadStringList(IDictionary d, string key)
        {
            var list = new List<string>();
            foreach (var item in Json.Json.GetArray(d, key))
            {
                var s = item as string;
                if (!string.IsNullOrEmpty(s)) list.Add(s);
            }
            return list;
        }

        private static int Clamp(int v, int min, int max)
        {
            return v < min ? min : (v > max ? max : v);
        }

        public void Save(string path)
        {
            File.WriteAllText(path, ToJson(), new UTF8Encoding(true));
        }

        public static RulesFile Load(string path)
        {
            return FromJson(File.ReadAllText(path, Encoding.UTF8));
        }

        /// <summary>האם מותר לגעת בפסקה הזו בכלל.</summary>
        public bool IsProtected(Model.ParagraphInfo p)
        {
            if (p == null) return true;

            foreach (var prefix in NeverTouchPrefixes)
            {
                if (!string.IsNullOrEmpty(prefix) &&
                    p.Text != null &&
                    p.Text.TrimStart().StartsWith(prefix, StringComparison.Ordinal))
                    return true;
            }

            foreach (var style in PreserveExistingStyles)
            {
                if (string.Equals(p.StyleName, style, StringComparison.Ordinal))
                    return true;
            }

            return false;
        }
    }
}
