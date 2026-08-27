using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Sadar.Core.Rules;

namespace Sadar.Core.Engine
{
    /// <summary>
    /// הרצת כלי שורת פקודה — משותף לכל המנועים שעובדים דרך מנוי.
    ///
    /// כלל ברזל: הפרומפט עובר ב-stdin ולעולם לא כארגומנט.
    /// שלד של ספר יכול להגיע לעשרות אלפי תווים, ושורת פקודה בווינדוס
    /// חסומה סביב 32K — ארגומנט ארוך מדי נכשל באמצע ובלי הסבר.
    /// </summary>
    public static class CliProcess
    {
        public static int Run(
            string exePath,
            string[] args,
            string stdin,
            int timeoutMs,
            out string stdout,
            out string stderr)
        {
            bool isCmd = exePath.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase) ||
                         exePath.EndsWith(".bat", StringComparison.OrdinalIgnoreCase);

            var psi = new ProcessStartInfo
            {
                FileName = isCmd ? "cmd.exe" : exePath,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
                WorkingDirectory = Path.GetTempPath()
            };

            var sb = new StringBuilder();
            if (isCmd) sb.Append("/c ").Append(Quote(exePath));
            foreach (string a in args) sb.Append(' ').Append(Quote(a));
            psi.Arguments = sb.ToString().TrimStart();

            // התהליך הבן לא אמור לרשת הקשר מהתהליך שקרא לו
            psi.EnvironmentVariables.Remove("CLAUDECODE");
            psi.EnvironmentVariables.Remove("CLAUDE_CODE_ENTRYPOINT");

            using (var proc = new Process())
            {
                proc.StartInfo = psi;

                var outBuf = new StringBuilder();
                var errBuf = new StringBuilder();
                proc.OutputDataReceived += (s, e) => { if (e.Data != null) outBuf.AppendLine(e.Data); };
                proc.ErrorDataReceived += (s, e) => { if (e.Data != null) errBuf.AppendLine(e.Data); };

                proc.Start();
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();

                if (stdin != null)
                {
                    // תהליך שנפל מוקדם סוגר את הצינור. השגיאה האמיתית
                    // תופיע ב-stderr, ולכן לא מסתירים אותה בחריגת כתיבה.
                    try
                    {
                        using (var w = new StreamWriter(proc.StandardInput.BaseStream, new UTF8Encoding(false)))
                        {
                            w.Write(stdin);
                            w.Flush();
                        }
                    }
                    catch (IOException) { }
                    catch (ObjectDisposedException) { }
                }
                else
                {
                    proc.StandardInput.Close();
                }

                if (!proc.WaitForExit(timeoutMs))
                {
                    try { proc.Kill(); } catch { }
                    stdout = outBuf.ToString();
                    stderr = errBuf.ToString();
                    throw new EngineException(
                        "הכלי לא השיב בתוך " + (timeoutMs / 1000) + " שניות. המסמך לא נגע.");
                }

                proc.WaitForExit();
                stdout = outBuf.ToString();
                stderr = errBuf.ToString();
                return proc.ExitCode;
            }
        }

        /// <summary>
        /// ציטוט לפי כללי שורת הפקודה של Windows: רק לוכסן אחורי
        /// שקודם לגרש דורש הכפלה. הכפלת כל הלוכסנים הורסת כל נתיב.
        /// </summary>
        public static string Quote(string s)
        {
            if (string.IsNullOrEmpty(s)) return "\"\"";

            var sb = new StringBuilder(s.Length + 8);
            sb.Append('"');

            int backslashes = 0;
            foreach (char c in s)
            {
                if (c == '\\') { backslashes++; continue; }
                if (c == '"') sb.Append('\\', backslashes * 2 + 1).Append('"');
                else { sb.Append('\\', backslashes); sb.Append(c); }
                backslashes = 0;
            }

            sb.Append('\\', backslashes * 2);
            sb.Append('"');
            return sb.ToString();
        }

        /// <summary>איתור כלי לפי רשימת מועמדים ואז לפי PATH.</summary>
        public static string Locate(IEnumerable<string> candidates, params string[] pathNames)
        {
            foreach (string p in candidates)
                if (!string.IsNullOrEmpty(p) && File.Exists(p)) return p;

            string path = Environment.GetEnvironmentVariable("PATH");
            if (string.IsNullOrEmpty(path)) return null;

            foreach (string dir in path.Split(';'))
            {
                if (string.IsNullOrEmpty(dir)) continue;
                foreach (string name in pathNames)
                {
                    try
                    {
                        string full = Path.Combine(dir.Trim(), name);
                        if (File.Exists(full)) return full;
                    }
                    catch { }
                }
            }
            return null;
        }

        /// <summary>מיקומי התקנה של כלי שהותקן דרך npm גלובלי.</summary>
        public static IEnumerable<string> NpmCandidates(string packagePath, string binName)
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            // הבינארי האמיתי עדיף על ה-.cmd העוטף: מעטפת cmd מוסיפה
            // שכבת פירוק ארגומנטים משלה
            yield return Path.Combine(appData, @"npm\node_modules\" + packagePath + @"\bin\" + binName + ".exe");
            yield return Path.Combine(appData, @"npm\" + binName + ".cmd");
            yield return Path.Combine(appData, @"npm\" + binName);
        }
    }

    /// <summary>
    /// בסיס למנוע שעובד דרך מנוי, בכלי שורת פקודה.
    ///
    /// הכלים האלה אינם חושפים הנחיית מערכת נפרדת, ולכן היא נכתבת
    /// בראש הפרומפט. זה פחות נקי מ-API, אבל זה מה שמאפשר להשתמש
    /// במנוי הקיים במקום לשלם שוב לפי שימוש.
    /// </summary>
    public abstract class CliEngineBase : IDecisionEngine
    {
        protected readonly EngineSettings Settings;
        protected readonly EngineInfo Info;
        protected readonly string ExePath;

        protected CliEngineBase(EngineSettings settings, EngineInfo info)
        {
            Settings = settings;
            Info = info;
            ExePath = !string.IsNullOrEmpty(settings.CliPath) && File.Exists(settings.CliPath)
                ? settings.CliPath
                : Locate();
        }

        public string Name { get { return Info.DisplayName + (ExePath == null ? " (לא נמצא)" : ""); } }
        public string Path { get { return ExePath; } }

        protected abstract string Locate();
        protected abstract string[] BuildArgs();
        protected abstract string ExtractText(string stdout);

        /// <summary>הפקודה שמתקינה את הכלי — מוצגת למשתמש כשהוא חסר.</summary>
        public string InstallHint { get { return Info.HowToConnect; } }

        public virtual string CheckAvailability()
        {
            if (string.IsNullOrEmpty(ExePath))
                return "לא נמצאה התקנה של " + Info.DisplayName + " במחשב." +
                       Environment.NewLine + Info.HowToConnect;

            try
            {
                string o, e;
                int exit = CliProcess.Run(ExePath, new[] { "--version" }, null, 25000, out o, out e);
                if (exit != 0)
                    return Info.DisplayName + " נמצא אך אינו מגיב כראוי. " + Trim(e, 200);
            }
            catch (Exception ex)
            {
                return "לא ניתן להריץ את " + Info.DisplayName + ": " + ex.Message;
            }

            return ProbeAuth();
        }

        /// <summary>בדיקת אימות בפועל — הרצה מינימלית.</summary>
        protected virtual string ProbeAuth()
        {
            try
            {
                string o, e;
                CliProcess.Run(ExePath, BuildArgs(), "החזר בדיוק את המילה: תקין", 60000, out o, out e);

                string combined = (o ?? "") + " " + (e ?? "");
                if (LooksUnauthenticated(combined))
                    return Info.DisplayName + " מותקן אך אינו מחובר לחשבון." +
                           Environment.NewLine + Info.HowToConnect;

                return null;
            }
            catch (Exception ex)
            {
                return "בדיקת החיבור נכשלה: " + ex.Message;
            }
        }

        protected static bool LooksUnauthenticated(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            string t = text.ToLowerInvariant();
            return t.Contains("not logged in") ||
                   t.Contains("not authenticated") ||
                   t.Contains("please log in") ||
                   t.Contains("please run /login") ||
                   t.Contains("no credentials") ||
                   t.Contains("unauthorized") ||
                   t.Contains("authentication required");
        }

        public string Decide(string skeletonJson, RulesFile rules)
        {
            if (string.IsNullOrEmpty(ExePath))
                throw new EngineException(Info.DisplayName + " אינו מותקן או שלא אותר במחשב.");

            string prompt = ClaudeCodeEngine.SystemPrompt +
                            Environment.NewLine + Environment.NewLine +
                            "--- השלד ---" + Environment.NewLine +
                            skeletonJson;

            string o, e;
            int exit = CliProcess.Run(ExePath, BuildArgs(), prompt,
                Settings.TimeoutSeconds * 1000, out o, out e);

            if (LooksUnauthenticated((o ?? "") + " " + (e ?? "")))
                throw new EngineException(
                    Info.DisplayName + " אינו מחובר לחשבון. " + Info.HowToConnect);

            if (ClaudeCodeEngine.LooksLikeFilter(e) || ClaudeCodeEngine.LooksLikeFilter(o))
                throw new EngineException(
                    "החיבור נחסם על ידי הסינון ברשת. נסו 'מצב מסונן' בהגדרות, או מנוע אחר.", true);

            string text = ExtractText(o);
            string json = text == null ? null : Json.Json.ExtractObject(text);

            if (json == null)
                throw new EngineException(
                    "לא התקבלה תשובה תקינה מ" + Info.DisplayName + ". קוד יציאה " + exit + ". " +
                    Trim(e ?? o, 300));

            return json;
        }

        protected static string Trim(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Trim();
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }
    }

    /// <summary>Google Gemini CLI, על חשבון גוגל של המשתמש.</summary>
    public sealed class GeminiCliEngine : CliEngineBase
    {
        public GeminiCliEngine(EngineSettings settings)
            : base(settings, EngineInfo.ByKind(EngineKind.GeminiCli)) { }

        protected override string Locate()
        {
            return CliProcess.Locate(
                CliProcess.NpmCandidates(@"@google\gemini-cli", "gemini"),
                "gemini.cmd", "gemini.exe", "gemini.ps1");
        }

        protected override string[] BuildArgs()
        {
            // הפרומפט נכנס ב-stdin. הדגלים כאן מבקשים מצב לא-אינטראקטיבי בלבד.
            var args = new List<string>();
            string model = Settings.EffectiveModel;
            if (!string.IsNullOrEmpty(model)) { args.Add("-m"); args.Add(model); }
            return args.ToArray();
        }

        protected override string ExtractText(string stdout) { return stdout; }
    }

    /// <summary>OpenAI Codex CLI, על חשבון ChatGPT של המשתמש.</summary>
    public sealed class CodexCliEngine : CliEngineBase
    {
        public CodexCliEngine(EngineSettings settings)
            : base(settings, EngineInfo.ByKind(EngineKind.CodexCli)) { }

        protected override string Locate()
        {
            return CliProcess.Locate(
                CliProcess.NpmCandidates(@"@openai\codex", "codex"),
                "codex.cmd", "codex.exe", "codex.ps1");
        }

        protected override string[] BuildArgs()
        {
            // exec = מצב לא-אינטראקטיבי. הפרומפט מגיע ב-stdin.
            var args = new List<string> { "exec" };
            string model = Settings.EffectiveModel;
            if (!string.IsNullOrEmpty(model)) { args.Add("-m"); args.Add(model); }
            return args.ToArray();
        }

        protected override string ExtractText(string stdout)
        {
            // codex exec מדפיס שורות סטטוס לפני התשובה.
            // חילוץ ה-JSON מטפל בזה, אבל מחזירים הכל כדי שלא נחתוך תשובה תקינה.
            return stdout;
        }
    }
}
