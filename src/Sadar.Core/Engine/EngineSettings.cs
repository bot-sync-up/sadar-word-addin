using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Sadar.Core.Engine
{
    /// <summary>סוג המנוע. שני מסלולים: מנוי דרך כלי שורת פקודה, או מפתח API.</summary>
    public enum EngineKind
    {
        ClaudeCli,
        GeminiCli,
        CodexCli,
        AnthropicApi,
        OpenAiApi,
        GeminiApi,
        OpenRouterApi
    }

    /// <summary>תיאור מנוע להצגה ולבחירה.</summary>
    public sealed class EngineInfo
    {
        public EngineKind Kind;
        public string Id;            // מזהה יציב לשורת הפקודה ולקובץ ההגדרות
        public string DisplayName;
        public string Vendor;
        public bool NeedsApiKey;
        public string DefaultModel;
        public string HowToConnect;

        public static readonly EngineInfo[] All =
        {
            new EngineInfo {
                Kind = EngineKind.ClaudeCli, Id = "claude-cli",
                DisplayName = "קלוד — דרך המנוי", Vendor = "Anthropic",
                NeedsApiKey = false, DefaultModel = "sonnet",
                HowToConnect = "npm install -g @anthropic-ai/claude-code  ואז  claude  ו-/login"
            },
            new EngineInfo {
                Kind = EngineKind.GeminiCli, Id = "gemini-cli",
                DisplayName = "ג'מיני — דרך המנוי", Vendor = "Google",
                NeedsApiKey = false, DefaultModel = "gemini-2.5-flash",
                HowToConnect = "npm install -g @google/gemini-cli  ואז  gemini  והתחברות עם חשבון גוגל"
            },
            new EngineInfo {
                Kind = EngineKind.CodexCli, Id = "codex-cli",
                DisplayName = "ChatGPT — דרך המנוי", Vendor = "OpenAI",
                NeedsApiKey = false, DefaultModel = "gpt-5",
                HowToConnect = "npm install -g @openai/codex  ואז  codex  והתחברות עם חשבון ChatGPT"
            },
            new EngineInfo {
                Kind = EngineKind.AnthropicApi, Id = "anthropic-api",
                DisplayName = "קלוד — מפתח API", Vendor = "Anthropic",
                NeedsApiKey = true, DefaultModel = "claude-opus-5",
                HowToConnect = "מפתח מ-console.anthropic.com"
            },
            new EngineInfo {
                Kind = EngineKind.OpenAiApi, Id = "openai-api",
                DisplayName = "ChatGPT — מפתח API", Vendor = "OpenAI",
                NeedsApiKey = true, DefaultModel = "gpt-5",
                HowToConnect = "מפתח מ-platform.openai.com"
            },
            new EngineInfo {
                Kind = EngineKind.GeminiApi, Id = "gemini-api",
                DisplayName = "ג'מיני — מפתח API", Vendor = "Google",
                NeedsApiKey = true, DefaultModel = "gemini-2.5-flash",
                HowToConnect = "מפתח מ-aistudio.google.com"
            },
            new EngineInfo {
                Kind = EngineKind.OpenRouterApi, Id = "openrouter-api",
                DisplayName = "OpenRouter — מפתח API", Vendor = "OpenRouter",
                NeedsApiKey = true, DefaultModel = "anthropic/claude-sonnet-4.5",
                HowToConnect = "מפתח מ-openrouter.ai — שער אחד לעשרות מודלים"
            }
        };

        public static EngineInfo ById(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            foreach (var e in All)
                if (string.Equals(e.Id, id, StringComparison.OrdinalIgnoreCase)) return e;
            return null;
        }

        public static EngineInfo ByKind(EngineKind kind)
        {
            foreach (var e in All) if (e.Kind == kind) return e;
            return All[0];
        }
    }

    /// <summary>
    /// בחירת המנוע והחיבור אליו.
    ///
    /// נפרד מקובץ הכללים בכוונה: הכללים שייכים לסדרת הספרים ואפשר לשתף אותם,
    /// והחיבור שייך למחשב ומכיל סודות.
    /// </summary>
    public sealed class EngineSettings
    {
        public string EngineId = "claude-cli";
        public string Model = "";              // ריק = ברירת המחדל של המנוע
        public int TimeoutSeconds = 180;

        /// <summary>נתיב ידני לכלי שורת הפקודה, אם האיתור האוטומטי נכשל.</summary>
        public string CliPath = "";

        /// <summary>כתובת בסיס חלופית לשירותים תואמי-OpenAI. ריק = ברירת המחדל.</summary>
        public string BaseUrl = "";

        /// <summary>מפתחות מוצפנים, לפי מזהה מנוע. הערך הגלוי לעולם אינו נשמר.</summary>
        private readonly Dictionary<string, string> _protectedKeys =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public EngineInfo Info
        {
            get { return EngineInfo.ById(EngineId) ?? EngineInfo.All[0]; }
        }

        public string EffectiveModel
        {
            get { return string.IsNullOrEmpty(Model) ? Info.DefaultModel : Model; }
        }

        // ---------- מפתחות ----------

        public bool HasKey(string engineId)
        {
            string v;
            return _protectedKeys.TryGetValue(engineId, out v) && !string.IsNullOrEmpty(v);
        }

        public void SetKey(string engineId, string plainKey)
        {
            if (string.IsNullOrEmpty(plainKey)) _protectedKeys.Remove(engineId);
            else _protectedKeys[engineId] = SecretStore.Protect(plainKey);
        }

        /// <summary>
        /// מחזיר את המפתח הגלוי. נקרא רק ברגע הקריאה לשירות ואינו נשמר בשום מקום.
        /// </summary>
        public string GetKey(string engineId)
        {
            string v;
            if (!_protectedKeys.TryGetValue(engineId, out v) || string.IsNullOrEmpty(v)) return null;
            return SecretStore.Unprotect(v);
        }

        public string CurrentKey { get { return GetKey(EngineId); } }

        /// <summary>ארבעת התווים האחרונים בלבד, להצגה בממשק.</summary>
        public string KeyHint(string engineId)
        {
            string k = GetKey(engineId);
            if (string.IsNullOrEmpty(k)) return null;
            return k.Length <= 4 ? "····" : "····" + k.Substring(k.Length - 4);
        }

        // ---------- שמירה וטעינה ----------

        public string ToJson()
        {
            var w = new Json.Json.Writer();
            w.StartObject();
            w.Prop("engineId", EngineId);
            w.Prop("model", Model ?? "");
            w.Prop("timeoutSeconds", TimeoutSeconds);
            w.Prop("cliPath", CliPath ?? "");
            w.Prop("baseUrl", BaseUrl ?? "");

            w.Name("keys").StartObject();
            foreach (var kv in _protectedKeys) w.Prop(kv.Key, kv.Value);
            w.EndObject();

            w.EndObject();
            return w.ToString();
        }

        public static EngineSettings FromJson(string json)
        {
            var d = Json.Json.ParseObject(json);
            var s = new EngineSettings();

            s.EngineId = Json.Json.GetString(d, "engineId", s.EngineId);
            s.Model = Json.Json.GetString(d, "model", "");
            s.TimeoutSeconds = Clamp(Json.Json.GetInt(d, "timeoutSeconds", 180), 30, 900);
            s.CliPath = Json.Json.GetString(d, "cliPath", "");
            s.BaseUrl = Json.Json.GetString(d, "baseUrl", "");

            object keys;
            if (Json.Json.TryGet(d, "keys", out keys))
            {
                var map = keys as IDictionary;
                if (map != null)
                    foreach (DictionaryEntry e in map)
                    {
                        string k = Convert.ToString(e.Key);
                        string v = Convert.ToString(e.Value);
                        if (!string.IsNullOrEmpty(k) && !string.IsNullOrEmpty(v))
                            s._protectedKeys[k] = v;
                    }
            }

            if (EngineInfo.ById(s.EngineId) == null) s.EngineId = "claude-cli";
            return s;
        }

        private static int Clamp(int v, int min, int max)
        {
            return v < min ? min : (v > max ? max : v);
        }

        /// <summary>מיקום קובץ ההגדרות. משותף לתוסף ולכלי שורת הפקודה.</summary>
        public static string DefaultPath
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Sadar");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "מנוע.json");
            }
        }

        public static EngineSettings LoadDefault()
        {
            try
            {
                if (File.Exists(DefaultPath)) return Load(DefaultPath);
            }
            catch { /* קובץ פגום — לא חוסמים את המשתמש */ }
            return new EngineSettings();
        }

        public void SaveDefault() { Save(DefaultPath); }

        public void Save(string path)
        {
            File.WriteAllText(path, ToJson(), new UTF8Encoding(true));
        }

        public static EngineSettings Load(string path)
        {
            return FromJson(File.ReadAllText(path, Encoding.UTF8));
        }
    }

    /// <summary>
    /// הצפנת מפתחות API.
    ///
    /// משתמש ב-DPAPI של Windows בהיקף המשתמש: הקובץ שנשמר אינו קריא
    /// למשתמש אחר במחשב ואינו שמיש אם מעתיקים אותו למחשב אחר.
    /// זה לא כספת — אבל זה מונע מפתח בטקסט גלוי בתיקיית ההגדרות.
    /// </summary>
    public static class SecretStore
    {
        private static readonly byte[] Entropy =
            Encoding.UTF8.GetBytes("Sadar.EngineKeys.v1");

        public static string Protect(string plain)
        {
            if (string.IsNullOrEmpty(plain)) return "";
            try
            {
                byte[] data = Encoding.UTF8.GetBytes(plain);
                byte[] enc = ProtectedData.Protect(data, Entropy, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(enc);
            }
            catch
            {
                // אם DPAPI אינו זמין עדיף להיכשל מאשר לשמור מפתח גלוי
                return "";
            }
        }

        public static string Unprotect(string protectedValue)
        {
            if (string.IsNullOrEmpty(protectedValue)) return null;
            try
            {
                byte[] enc = Convert.FromBase64String(protectedValue);
                byte[] data = ProtectedData.Unprotect(enc, Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(data);
            }
            catch
            {
                // הקובץ הועתק ממחשב אחר או מפרופיל אחר
                return null;
            }
        }
    }
}
