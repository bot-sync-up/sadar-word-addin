using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Web.Script.Serialization;

namespace Sadar.Core.Json
{
    /// <summary>
    /// כתיבה וקריאה של JSON ללא תלות חיצונית.
    /// הכתיבה קומפקטית בכוונה — כל תו שנחסך הוא טוקן שנחסך בקריאה למודל.
    /// </summary>
    public static class Json
    {
        // ---------- כתיבה ----------

        public sealed class Writer
        {
            private readonly StringBuilder _sb = new StringBuilder(4096);
            private bool _needComma;

            public Writer StartObject() { Sep(); _sb.Append('{'); _needComma = false; return this; }
            public Writer EndObject() { _sb.Append('}'); _needComma = true; return this; }
            public Writer StartArray() { Sep(); _sb.Append('['); _needComma = false; return this; }
            public Writer EndArray() { _sb.Append(']'); _needComma = true; return this; }

            public Writer Name(string name)
            {
                Sep();
                AppendString(name);
                _sb.Append(':');
                _needComma = false;
                return this;
            }

            public Writer Value(string v)
            {
                Sep();
                if (v == null) _sb.Append("null"); else AppendString(v);
                _needComma = true;
                return this;
            }

            public Writer Value(int v)
            {
                Sep();
                _sb.Append(v.ToString(CultureInfo.InvariantCulture));
                _needComma = true;
                return this;
            }

            public Writer Value(double v)
            {
                Sep();
                _sb.Append(v.ToString("0.##", CultureInfo.InvariantCulture));
                _needComma = true;
                return this;
            }

            public Writer Value(bool v)
            {
                Sep();
                _sb.Append(v ? "true" : "false");
                _needComma = true;
                return this;
            }

            public Writer Prop(string name, string v) { Name(name); Value(v); return this; }
            public Writer Prop(string name, int v) { Name(name); Value(v); return this; }
            public Writer Prop(string name, double v) { Name(name); Value(v); return this; }
            public Writer Prop(string name, bool v) { Name(name); Value(v); return this; }

            private void Sep()
            {
                if (_needComma) { _sb.Append(','); _needComma = false; }
            }

            private void AppendString(string s)
            {
                _sb.Append('"');
                foreach (char c in s)
                {
                    switch (c)
                    {
                        case '"': _sb.Append("\\\""); break;
                        case '\\': _sb.Append("\\\\"); break;
                        case '\n': _sb.Append("\\n"); break;
                        case '\r': _sb.Append("\\r"); break;
                        case '\t': _sb.Append("\\t"); break;
                        case '\b': _sb.Append("\\b"); break;
                        case '\f': _sb.Append("\\f"); break;
                        default:
                            // עברית נשלחת כתווים ממשיים ולא כ-\uXXXX: חוסך כשלושה רבעים מהנפח
                            if (c < 0x20) _sb.Append("\\u").Append(((int)c).ToString("x4"));
                            else _sb.Append(c);
                            break;
                    }
                }
                _sb.Append('"');
            }

            public override string ToString() { return _sb.ToString(); }
            public int Length { get { return _sb.Length; } }
        }

        // ---------- קריאה ----------

        private static readonly JavaScriptSerializer Serializer = CreateSerializer();

        private static JavaScriptSerializer CreateSerializer()
        {
            var s = new JavaScriptSerializer();
            s.MaxJsonLength = int.MaxValue;
            s.RecursionLimit = 200;
            return s;
        }

        /// <summary>מפרק JSON למילון/רשימות. זורק חריגה על קלט לא תקין.</summary>
        public static object Parse(string json)
        {
            if (string.IsNullOrEmpty(json)) throw new FormatException("JSON ריק");
            return Serializer.DeserializeObject(json);
        }

        public static Dictionary<string, object> ParseObject(string json)
        {
            var o = Parse(json) as Dictionary<string, object>;
            if (o == null) throw new FormatException("צפוי אובייקט JSON ברמה העליונה");
            return o;
        }

        // ---------- גישה בטוחה לערכים ----------

        public static bool TryGet(IDictionary d, string key, out object value)
        {
            value = null;
            if (d == null || !d.Contains(key)) return false;
            value = d[key];
            return value != null;
        }

        public static string GetString(IDictionary d, string key, string fallback = null)
        {
            object v;
            if (!TryGet(d, key, out v)) return fallback;
            return v as string ?? Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        public static int GetInt(IDictionary d, string key, int fallback = 0)
        {
            object v;
            if (!TryGet(d, key, out v)) return fallback;
            try { return Convert.ToInt32(v, CultureInfo.InvariantCulture); }
            catch { return fallback; }
        }

        public static bool GetBool(IDictionary d, string key, bool fallback = false)
        {
            object v;
            if (!TryGet(d, key, out v)) return fallback;
            if (v is bool) return (bool)v;
            bool b;
            return bool.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), out b) ? b : fallback;
        }

        public static IEnumerable<object> GetArray(IDictionary d, string key)
        {
            object v;
            if (!TryGet(d, key, out v)) return new object[0];
            var arr = v as IEnumerable;
            if (arr == null || v is string) return new object[0];
            var list = new List<object>();
            foreach (var item in arr) list.Add(item);
            return list;
        }

        /// <summary>
        /// חילוץ אובייקט JSON מתוך תשובה שעטופה בטקסט או בגדרות markdown.
        /// המודל אמור להחזיר JSON נקי; זו רשת ביטחון ולא היתר לפלט מלוכלך.
        /// </summary>
        public static string ExtractObject(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return null;

            int fence = raw.IndexOf("```", StringComparison.Ordinal);
            if (fence >= 0)
            {
                int start = raw.IndexOf('\n', fence);
                int end = raw.IndexOf("```", fence + 3, StringComparison.Ordinal);
                if (start > 0 && end > start) raw = raw.Substring(start + 1, end - start - 1);
            }

            int open = raw.IndexOf('{');
            if (open < 0) return null;

            int depth = 0;
            bool inString = false, escape = false;
            for (int i = open; i < raw.Length; i++)
            {
                char c = raw[i];
                if (escape) { escape = false; continue; }
                if (c == '\\' && inString) { escape = true; continue; }
                if (c == '"') { inString = !inString; continue; }
                if (inString) continue;
                if (c == '{') depth++;
                else if (c == '}')
                {
                    depth--;
                    if (depth == 0) return raw.Substring(open, i - open + 1);
                }
            }
            return null;
        }
    }
}
