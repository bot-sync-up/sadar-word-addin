using System;
using System.Collections.Generic;

namespace Sadar.Core.Engine
{
    /// <summary>
    /// בונה את המנוע לפי ההגדרות.
    ///
    /// כל הצינור שמעל מדבר עם IDecisionEngine בלבד ואינו יודע איזה ספק
    /// נבחר. זו הסיבה שהוספת מנוע חדש נוגעת רק בקבצים שבתיקייה הזו.
    /// </summary>
    public static class EngineFactory
    {
        public static IDecisionEngine Create(EngineSettings settings)
        {
            if (settings == null) settings = new EngineSettings();

            switch (settings.Info.Kind)
            {
                case EngineKind.ClaudeCli:
                    return new ClaudeCodeEngine(
                        string.IsNullOrEmpty(settings.CliPath) ? null : settings.CliPath,
                        settings.EffectiveModel,
                        settings.TimeoutSeconds);

                case EngineKind.GeminiCli:
                    return new GeminiCliEngine(settings);

                case EngineKind.CodexCli:
                    return new CodexCliEngine(settings);

                case EngineKind.AnthropicApi:
                    return new AnthropicApiEngine(settings);

                case EngineKind.OpenAiApi:
                    return new OpenAiApiEngine(settings, EngineKind.OpenAiApi);

                case EngineKind.OpenRouterApi:
                    return new OpenAiApiEngine(settings, EngineKind.OpenRouterApi);

                case EngineKind.GeminiApi:
                    return new GeminiApiEngine(settings);

                default:
                    throw new EngineException("מנוע לא מוכר: " + settings.EngineId);
            }
        }

        /// <summary>בונה מנוע לפי מזהה, בלי לגעת בהגדרות השמורות. לשימוש שורת הפקודה.</summary>
        public static IDecisionEngine CreateById(string engineId, EngineSettings baseSettings)
        {
            var info = EngineInfo.ById(engineId);
            if (info == null)
                throw new EngineException(
                    "מנוע לא מוכר: " + engineId + ". הרשימה המלאה: sadar engines");

            var s = baseSettings ?? new EngineSettings();
            s.EngineId = info.Id;
            return Create(s);
        }

        /// <summary>סטטוס של כל המנועים — למסך ההגדרות ול-doctor.</summary>
        public sealed class EngineStatus
        {
            public EngineInfo Info;
            public bool Configured;   // מותקן, או שיש לו מפתח
            public string Problem;    // null = מוכן לשימוש
            public string Detail;     // נתיב הכלי, או רמז למפתח

            public bool Ready { get { return Problem == null; } }
        }

        public static List<EngineStatus> Survey(EngineSettings settings, bool deepCheck)
        {
            var list = new List<EngineStatus>();

            foreach (var info in EngineInfo.All)
            {
                var status = new EngineStatus { Info = info };

                var probe = new EngineSettings
                {
                    EngineId = info.Id,
                    Model = "",
                    TimeoutSeconds = 60,
                    CliPath = info.Id == settings.EngineId ? settings.CliPath : "",
                    BaseUrl = info.Id == settings.EngineId ? settings.BaseUrl : ""
                };

                if (info.NeedsApiKey)
                {
                    string key = settings.GetKey(info.Id);
                    status.Configured = !string.IsNullOrEmpty(key);
                    status.Detail = settings.KeyHint(info.Id);

                    if (!status.Configured)
                    {
                        status.Problem = "לא הוגדר מפתח";
                        list.Add(status);
                        continue;
                    }
                    probe.SetKey(info.Id, key);
                }

                try
                {
                    var engine = Create(probe);

                    var cli = engine as CliEngineBase;
                    if (cli != null)
                    {
                        status.Detail = cli.Path;
                        status.Configured = !string.IsNullOrEmpty(cli.Path);
                    }

                    var claude = engine as ClaudeCodeEngine;
                    if (claude != null)
                    {
                        status.Detail = claude.ExePath;
                        status.Configured = !string.IsNullOrEmpty(claude.ExePath);
                    }

                    if (!status.Configured && !info.NeedsApiKey)
                    {
                        status.Problem = "לא מותקן";
                    }
                    else if (deepCheck)
                    {
                        status.Problem = engine.CheckAvailability();
                    }
                    else
                    {
                        status.Problem = null;
                    }
                }
                catch (Exception ex)
                {
                    status.Problem = ex.Message;
                }

                list.Add(status);
            }

            return list;
        }
    }
}
