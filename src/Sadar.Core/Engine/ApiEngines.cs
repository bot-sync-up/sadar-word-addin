using System;
using System.Net;

namespace Sadar.Core.Engine
{
    /// <summary>
    /// Anthropic Messages API.
    /// POST https://api.anthropic.com/v1/messages
    /// </summary>
    public sealed class AnthropicApiEngine : HttpEngineBase
    {
        /// <summary>גרסת ה-API. נדרשת בכל בקשה.</summary>
        private const string ApiVersion = "2023-06-01";

        public AnthropicApiEngine(EngineSettings settings)
            : base(settings, EngineInfo.ByKind(EngineKind.AnthropicApi)) { }

        public override string Endpoint
        {
            get
            {
                string b = string.IsNullOrEmpty(Settings.BaseUrl)
                    ? "https://api.anthropic.com" : Settings.BaseUrl.TrimEnd('/');
                return b + "/v1/messages";
            }
        }

        protected override void AddHeaders(HttpWebRequest request, string apiKey)
        {
            request.Headers["x-api-key"] = apiKey;
            request.Headers["anthropic-version"] = ApiVersion;
        }

        public override string BuildRequestBody(string systemPrompt, string userPrompt, string model)
        {
            var w = new Json.Json.Writer();
            w.StartObject();
            w.Prop("model", model);
            w.Prop("max_tokens", 16000);
            w.Prop("system", systemPrompt);
            w.Name("messages").StartArray();
            w.StartObject().Prop("role", "user").Prop("content", userPrompt).EndObject();
            w.EndArray();
            w.EndObject();
            return w.ToString();
        }

        public override string ExtractText(string responseJson)
        {
            var root = Json.Json.ParseObject(responseJson);

            // stop_reason "refusal" מגיע כ-HTTP 200 ותוכן ריק.
            // בלי הבדיקה הזו זה נראה כמו תשובה פגומה במקום סירוב.
            string stop = Json.Json.GetString(root, "stop_reason", null);
            if (string.Equals(stop, "refusal", StringComparison.Ordinal))
                throw new EngineException(
                    "המודל סירב לענות על הבקשה. נסו לנסח מחדש את ההגדרה המבנית בכללים.");

            var sb = new System.Text.StringBuilder();
            foreach (var block in Json.Json.GetArray(root, "content"))
            {
                var d = block as System.Collections.IDictionary;
                if (d == null) continue;
                if (!string.Equals(Json.Json.GetString(d, "type", ""), "text", StringComparison.Ordinal))
                    continue;
                sb.Append(Json.Json.GetString(d, "text", ""));
            }
            return sb.ToString();
        }
    }

    /// <summary>
    /// OpenAI Chat Completions, ותואמיו.
    ///
    /// אותו מימוש משרת גם את OpenRouter וגם כל שירות תואם-OpenAI אחר,
    /// כי כולם חולקים את אותו פורמט. מה שמשתנה הוא כתובת הבסיס בלבד.
    /// </summary>
    public sealed class OpenAiApiEngine : HttpEngineBase
    {
        private readonly string _defaultBase;

        public OpenAiApiEngine(EngineSettings settings, EngineKind kind)
            : base(settings, EngineInfo.ByKind(kind))
        {
            _defaultBase = kind == EngineKind.OpenRouterApi
                ? "https://openrouter.ai/api"
                : "https://api.openai.com";
        }

        public override string Endpoint
        {
            get
            {
                string b = string.IsNullOrEmpty(Settings.BaseUrl)
                    ? _defaultBase : Settings.BaseUrl.TrimEnd('/');
                return b + "/v1/chat/completions";
            }
        }

        protected override void AddHeaders(HttpWebRequest request, string apiKey)
        {
            request.Headers["Authorization"] = "Bearer " + apiKey;

            if (Info.Kind == EngineKind.OpenRouterApi)
            {
                // OpenRouter מבקש לזהות את האפליקציה הקוראת
                request.Headers["HTTP-Referer"] = "https://github.com/bot-sync-up/sadar-word-addin";
                request.Headers["X-Title"] = "Sadar";
            }
        }

        public override string BuildRequestBody(string systemPrompt, string userPrompt, string model)
        {
            var w = new Json.Json.Writer();
            w.StartObject();
            w.Prop("model", model);

            w.Name("messages").StartArray();
            w.StartObject().Prop("role", "system").Prop("content", systemPrompt).EndObject();
            w.StartObject().Prop("role", "user").Prop("content", userPrompt).EndObject();
            w.EndArray();

            // מבקשים JSON ברמת ה-API ולא רק בהנחיה. לא כל מודל תומך,
            // ולכן ההנחיה נשארת גם היא — שתי חגורות במקום אחת.
            w.Name("response_format").StartObject().Prop("type", "json_object").EndObject();

            w.EndObject();
            return w.ToString();
        }

        public override string ExtractText(string responseJson)
        {
            var root = Json.Json.ParseObject(responseJson);

            foreach (var choice in Json.Json.GetArray(root, "choices"))
            {
                var c = choice as System.Collections.IDictionary;
                if (c == null) continue;

                object msg;
                if (!Json.Json.TryGet(c, "message", out msg)) continue;
                var m = msg as System.Collections.IDictionary;
                if (m == null) continue;

                return Json.Json.GetString(m, "content", "");
            }
            return "";
        }
    }

    /// <summary>
    /// Google Gemini — generateContent.
    /// המפתח עובר בכותרת ולא ב-query string, כדי שלא יופיע ביומני שרתים ובפרוקסי.
    /// </summary>
    public sealed class GeminiApiEngine : HttpEngineBase
    {
        public GeminiApiEngine(EngineSettings settings)
            : base(settings, EngineInfo.ByKind(EngineKind.GeminiApi)) { }

        public override string Endpoint
        {
            get
            {
                string b = string.IsNullOrEmpty(Settings.BaseUrl)
                    ? "https://generativelanguage.googleapis.com" : Settings.BaseUrl.TrimEnd('/');
                return b + "/v1beta/models/" + Settings.EffectiveModel + ":generateContent";
            }
        }

        protected override void AddHeaders(HttpWebRequest request, string apiKey)
        {
            request.Headers["x-goog-api-key"] = apiKey;
        }

        public override string BuildRequestBody(string systemPrompt, string userPrompt, string model)
        {
            var w = new Json.Json.Writer();
            w.StartObject();

            w.Name("systemInstruction").StartObject()
                .Name("parts").StartArray()
                .StartObject().Prop("text", systemPrompt).EndObject()
                .EndArray()
                .EndObject();

            w.Name("contents").StartArray()
                .StartObject()
                .Prop("role", "user")
                .Name("parts").StartArray()
                .StartObject().Prop("text", userPrompt).EndObject()
                .EndArray()
                .EndObject()
                .EndArray();

            w.Name("generationConfig").StartObject()
                .Prop("responseMimeType", "application/json")
                .Prop("temperature", 0)
                .EndObject();

            w.EndObject();
            return w.ToString();
        }

        public override string ExtractText(string responseJson)
        {
            var root = Json.Json.ParseObject(responseJson);

            foreach (var cand in Json.Json.GetArray(root, "candidates"))
            {
                var c = cand as System.Collections.IDictionary;
                if (c == null) continue;

                // חסימת בטיחות מגיעה כ-HTTP 200 בלי תוכן
                string finish = Json.Json.GetString(c, "finishReason", "");
                if (string.Equals(finish, "SAFETY", StringComparison.OrdinalIgnoreCase))
                    throw new EngineException(
                        "ג'מיני חסם את הבקשה מטעמי בטיחות. נסו מנוע אחר, או 'מצב מסונן' " +
                        "שאינו שולח טקסט כלל.");

                object content;
                if (!Json.Json.TryGet(c, "content", out content)) continue;
                var cd = content as System.Collections.IDictionary;
                if (cd == null) continue;

                var sb = new System.Text.StringBuilder();
                foreach (var part in Json.Json.GetArray(cd, "parts"))
                {
                    var p = part as System.Collections.IDictionary;
                    if (p == null) continue;
                    sb.Append(Json.Json.GetString(p, "text", ""));
                }
                return sb.ToString();
            }

            // בקשה שנחסמה כולה מדווחת ב-promptFeedback
            object feedback;
            if (Json.Json.TryGet(root, "promptFeedback", out feedback))
            {
                var f = feedback as System.Collections.IDictionary;
                string blocked = f == null ? null : Json.Json.GetString(f, "blockReason", null);
                if (!string.IsNullOrEmpty(blocked))
                    throw new EngineException("ג'מיני חסם את הבקשה: " + blocked);
            }

            return "";
        }
    }
}
