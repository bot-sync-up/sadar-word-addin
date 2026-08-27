using System;
using System.Collections.Generic;
using Sadar.Core.Engine;

namespace Sadar.Cli
{
    /// <summary>
    /// בדיקות שכבת המנועים.
    ///
    /// כולן רצות בלי רשת ובלי מפתח אמיתי: מה שנבדק הוא חוזה החיווט מול
    /// כל ספק — לאן פונים, מה נשלח, איך מפענחים את התשובה, ומה קורה
    /// כשהספק מסרב או חוסם.
    /// </summary>
    public static class QaEngineTests
    {
        public static void Run(Action<string> section, Action<string, bool, string> check)
        {
            section("מנועים — רישום והגדרות");
            AllEnginesResolve(check);
            SettingsRoundTrip(check);
            SecretRoundTrip(check);
            SecretRejectsGarbage(check);
            UnknownEngineRejected(check);
            CorruptSettingsFallBack(check);

            section("מנועים — חוזה החיווט");
            AnthropicWireFormat(check);
            OpenAiWireFormat(check);
            GeminiWireFormat(check);
            OpenRouterUsesOwnHost(check);
            CustomBaseUrlHonoured(check);

            section("מנועים — סירוב וחסימה");
            RefusalIsReported(check);
            GeminiSafetyBlockIsReported(check);
        }

        // ---------- רישום והגדרות ----------

        private static void AllEnginesResolve(Action<string, bool, string> check)
        {
            int built = 0;
            string failure = null;

            foreach (var info in EngineInfo.All)
            {
                var s = new EngineSettings { EngineId = info.Id };
                if (info.NeedsApiKey) s.SetKey(info.Id, "test-key");

                try
                {
                    var engine = EngineFactory.Create(s);
                    if (engine != null && !string.IsNullOrEmpty(engine.Name)) built++;
                }
                catch (Exception ex) { failure = info.Id + ": " + ex.Message; break; }
            }

            check("כל המנועים נבנים", built == EngineInfo.All.Length,
                failure ?? (built + " מתוך " + EngineInfo.All.Length));

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool unique = true;
            foreach (var i in EngineInfo.All) if (!seen.Add(i.Id)) unique = false;
            check("לכל מנוע מזהה ייחודי", unique, "");

            bool haveModels = true;
            foreach (var i in EngineInfo.All)
                if (string.IsNullOrEmpty(i.DefaultModel)) haveModels = false;
            check("לכל מנוע מודל ברירת מחדל", haveModels, "");
        }

        private static void SettingsRoundTrip(Action<string, bool, string> check)
        {
            var s = new EngineSettings
            {
                EngineId = "gemini-api",
                Model = "gemini-2.5-pro",
                TimeoutSeconds = 240,
                BaseUrl = "https://example.test/v1"
            };
            s.SetKey("gemini-api", "AIza-secret-key-1234");
            s.SetKey("openai-api", "sk-other-9999");

            string json = s.ToJson();
            var back = EngineSettings.FromJson(json);

            check("המנוע נשמר", back.EngineId == "gemini-api", "");
            check("המודל נשמר", back.Model == "gemini-2.5-pro", "");
            check("הזמן הקצוב נשמר", back.TimeoutSeconds == 240, "");
            check("כתובת הבסיס נשמרת", back.BaseUrl == "https://example.test/v1", "");
            check("המפתח שרד סבב שמירה וטעינה",
                back.GetKey("gemini-api") == "AIza-secret-key-1234", "");
            check("מפתח של ספק אחר נשמר בנפרד",
                back.GetKey("openai-api") == "sk-other-9999", "");
            check("מפתח שלא הוגדר מחזיר ריק", back.GetKey("anthropic-api") == null, "");

            // המבחן החשוב: הקובץ שנשמר לדיסק אינו מכיל את המפתח בטקסט גלוי
            check("המפתח אינו מופיע גלוי בקובץ",
                !json.Contains("AIza-secret-key-1234"), "");
            check("הרמז חושף רק את ארבעת התווים האחרונים",
                back.KeyHint("gemini-api") == "····1234", back.KeyHint("gemini-api"));
        }

        private static void SecretRoundTrip(Action<string, bool, string> check)
        {
            const string secret = "sk-ant-api03-0123456789-אבג";
            string prot = SecretStore.Protect(secret);

            check("ההצפנה מחזירה ערך", !string.IsNullOrEmpty(prot), "");
            check("הערך המוצפן שונה מהמקור", prot != secret, "");
            check("הפענוח מחזיר בדיוק את המקור", SecretStore.Unprotect(prot) == secret, "");
        }

        private static void SecretRejectsGarbage(Action<string, bool, string> check)
        {
            check("פענוח של טקסט שאינו base64 מחזיר ריק ולא קורס",
                SecretStore.Unprotect("לא-base64-בכלל") == null, "");
            check("פענוח של base64 שאינו שלנו מחזיר ריק",
                SecretStore.Unprotect("QUJDREVGRw==") == null, "");
        }

        private static void UnknownEngineRejected(Action<string, bool, string> check)
        {
            bool threw = false;
            try { EngineFactory.CreateById("gpt-9-turbo-max", new EngineSettings()); }
            catch (EngineException) { threw = true; }
            check("מנוע לא מוכר נדחה בהודעה ברורה", threw, "");
        }

        private static void CorruptSettingsFallBack(Action<string, bool, string> check)
        {
            var s = EngineSettings.FromJson("{\"engineId\":\"אין-כזה\"}");
            check("מזהה מנוע לא חוקי חוזר לברירת מחדל", s.EngineId == "claude-cli", s.EngineId);

            var clamped = EngineSettings.FromJson("{\"timeoutSeconds\":999999}");
            check("זמן קצוב מוגזם נחתך", clamped.TimeoutSeconds <= 900,
                clamped.TimeoutSeconds.ToString());
        }

        // ---------- חוזה החיווט ----------

        private static HttpEngineBase Api(EngineKind kind, string model)
        {
            var s = new EngineSettings { EngineId = EngineInfo.ByKind(kind).Id, Model = model };
            s.SetKey(s.EngineId, "test-key");
            return (HttpEngineBase)EngineFactory.Create(s);
        }

        private static void AnthropicWireFormat(Action<string, bool, string> check)
        {
            var e = Api(EngineKind.AnthropicApi, "claude-opus-5");

            check("נקודת הקצה של Anthropic",
                e.Endpoint == "https://api.anthropic.com/v1/messages", e.Endpoint);

            string body = e.BuildRequestBody("SYS", "USER", "claude-opus-5");
            var d = Core.Json.Json.ParseObject(body);

            check("גוף הבקשה הוא JSON תקין", d != null, "");
            check("המודל נשלח", Core.Json.Json.GetString(d, "model", "") == "claude-opus-5", "");
            check("הנחיית המערכת בשדה system",
                Core.Json.Json.GetString(d, "system", "") == "SYS", "");
            check("max_tokens נשלח", Core.Json.Json.GetInt(d, "max_tokens", 0) > 0, "");
            check("הפרומפט נשלח כהודעת user", body.Contains("\"role\":\"user\""), "");

            string text = e.ExtractText(
                "{\"content\":[{\"type\":\"text\",\"text\":\"HELLO\"}]}");
            check("חילוץ התשובה", text == "HELLO", text);

            string multi = e.ExtractText(
                "{\"content\":[{\"type\":\"thinking\",\"thinking\":\"X\"}," +
                "{\"type\":\"text\",\"text\":\"A\"},{\"type\":\"text\",\"text\":\"B\"}]}");
            check("בלוקים שאינם טקסט מדולגים והשאר משורשר", multi == "AB", multi);
        }

        private static void OpenAiWireFormat(Action<string, bool, string> check)
        {
            var e = Api(EngineKind.OpenAiApi, "gpt-5");

            check("נקודת הקצה של OpenAI",
                e.Endpoint == "https://api.openai.com/v1/chat/completions", e.Endpoint);

            string body = e.BuildRequestBody("SYS", "USER", "gpt-5");
            check("גוף הבקשה הוא JSON תקין", Core.Json.Json.ParseObject(body) != null, "");
            check("הנחיית המערכת כהודעת system", body.Contains("\"role\":\"system\""), "");
            check("מבוקש פלט JSON ברמת ה-API", body.Contains("json_object"), "");

            string text = e.ExtractText(
                "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"HELLO\"}}]}");
            check("חילוץ התשובה", text == "HELLO", text);
        }

        private static void GeminiWireFormat(Action<string, bool, string> check)
        {
            var e = Api(EngineKind.GeminiApi, "gemini-2.5-flash");

            check("נקודת הקצה של ג'מיני כוללת את שם המודל",
                e.Endpoint.EndsWith("/v1beta/models/gemini-2.5-flash:generateContent"), e.Endpoint);

            string body = e.BuildRequestBody("SYS", "USER", "gemini-2.5-flash");
            check("גוף הבקשה הוא JSON תקין", Core.Json.Json.ParseObject(body) != null, "");
            check("הנחיית המערכת ב-systemInstruction", body.Contains("systemInstruction"), "");
            check("מבוקש פלט JSON", body.Contains("application/json"), "");

            string text = e.ExtractText(
                "{\"candidates\":[{\"content\":{\"parts\":[{\"text\":\"HELLO\"}]}}]}");
            check("חילוץ התשובה", text == "HELLO", text);
        }

        private static void OpenRouterUsesOwnHost(Action<string, bool, string> check)
        {
            var e = Api(EngineKind.OpenRouterApi, "anthropic/claude-sonnet-4.5");
            check("OpenRouter פונה לשרת שלו",
                e.Endpoint == "https://openrouter.ai/api/v1/chat/completions", e.Endpoint);
        }

        private static void CustomBaseUrlHonoured(Action<string, bool, string> check)
        {
            var s = new EngineSettings
            {
                EngineId = "openai-api",
                BaseUrl = "https://my-gateway.local/proxy/"
            };
            s.SetKey("openai-api", "k");

            var e = (HttpEngineBase)EngineFactory.Create(s);
            check("כתובת בסיס מותאמת מכובדת, בלי לוכסן כפול",
                e.Endpoint == "https://my-gateway.local/proxy/v1/chat/completions", e.Endpoint);
        }

        // ---------- סירוב וחסימה ----------

        private static void RefusalIsReported(Action<string, bool, string> check)
        {
            var e = Api(EngineKind.AnthropicApi, "claude-opus-5");

            bool threw = false;
            try { e.ExtractText("{\"stop_reason\":\"refusal\",\"content\":[]}"); }
            catch (EngineException) { threw = true; }

            check("סירוב מדווח כשגיאה ולא כתשובה ריקה", threw, "");
        }

        private static void GeminiSafetyBlockIsReported(Action<string, bool, string> check)
        {
            var e = Api(EngineKind.GeminiApi, "gemini-2.5-flash");

            bool safety = false;
            try { e.ExtractText("{\"candidates\":[{\"finishReason\":\"SAFETY\"}]}"); }
            catch (EngineException) { safety = true; }
            check("חסימת בטיחות מדווחת", safety, "");

            bool prompt = false;
            try { e.ExtractText("{\"promptFeedback\":{\"blockReason\":\"OTHER\"}}"); }
            catch (EngineException) { prompt = true; }
            check("חסימת בקשה מדווחת", prompt, "");
        }
    }
}
