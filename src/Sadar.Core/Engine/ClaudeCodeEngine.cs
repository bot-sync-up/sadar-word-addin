using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Sadar.Core.Rules;

namespace Sadar.Core.Engine
{
    /// <summary>
    /// גשר אל Claude Code המותקן והמחובר אצל המשתמש.
    ///
    /// אין כאן מפתח API ואין שרת. התוסף מריץ את ה-CLI של המשתמש
    /// על המחשב שלו, בדיוק כפי שהוא היה מריץ אותו בעצמו.
    ///
    /// הסוכן מורץ בלי אף כלי: הוא לא כותב קבצים, לא מריץ פקודות ולא ניגש לרשת.
    /// הוא מקבל טקסט ומחזיר JSON.
    /// </summary>
    public sealed class ClaudeCodeEngine : IDecisionEngine
    {
        public const string SystemPrompt =
            "אתה מנוע החלטות עימוד עבור מסמכי וורד בעברית. " +
            "אתה מקבל שלד מבני של מסמך — אינדקסי פסקאות, תחילית קצרה ומאפייני עיצוב — " +
            "ומחזיר אך ורק אובייקט JSON יחיד של החלטות סגנון.\n\n" +
            "כללים מוחלטים:\n" +
            "1. אל תחזיר טקסט מהמסמך. אל תצטט, אל תשכתב, אל תתקן ואל תסכם.\n" +
            "2. אל תחזיר הסבר, הקדמה, סיכום או גדרות markdown. רק JSON.\n" +
            "3. השתמש אך ורק בשמות הסגנונות שברשימת allowedStyles.\n" +
            "4. deletes מותר אך ורק על פסקאות שסומנו empty:true.\n" +
            "5. פסקה שאינך בטוח לגביה — שים אותה ב-uncertain עם סיבה קצרה, ואל תנחש.\n\n" +
            "פורמט התשובה:\n" +
            "{\"decisions\":[{\"i\":<מספר>,\"style\":\"<שם סגנון>\"}]," +
            "\"deletes\":[<מספרים>]," +
            "\"uncertain\":[{\"i\":<מספר>,\"style\":\"<שם סגנון>\",\"why\":\"<סיבה קצרה>\"}]}";

        private readonly string _exePath;
        private readonly string _model;
        private readonly int _timeoutMs;

        public ClaudeCodeEngine(string exePath = null, string model = "sonnet", int timeoutSeconds = 180)
        {
            _exePath = exePath ?? Locate();
            _model = model;
            _timeoutMs = timeoutSeconds * 1000;
        }

        public string Name
        {
            get { return "Claude Code" + (_exePath == null ? " (לא נמצא)" : ""); }
        }

        public string ExePath { get { return _exePath; } }

        // ---------- איתור ----------

        /// <summary>
        /// מיקומי ההתקנה הידועים, לפי סדר עדיפות.
        /// הנתיב של אפליקציית הדסקטופ נבדק אחרון: הבינארי קיים שם,
        /// אך הוא נשען על אימות שהאפליקציה מספקת ולא תמיד זמין להרצה עצמאית.
        /// </summary>
        public static IEnumerable<string> CandidatePaths()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            // הבינארי האמיתי של ההתקנה דרך npm. עדיף עליו על פני ה-.cmd העוטף:
            // מעטפת cmd מוסיפה שכבת פירוק ארגומנטים משלה, ופרומפט עם גרשיים נשבר בה.
            yield return Path.Combine(appData,
                @"npm\node_modules\@anthropic-ai\claude-code\bin\claude.exe");

            yield return Path.Combine(appData, @"npm\claude.cmd");
            yield return Path.Combine(localApp, @"Programs\claude\claude.exe");
            yield return Path.Combine(profile, @".local\bin\claude.exe");
            yield return Path.Combine(profile, @".claude\local\claude.exe");

            // הבינארי הארוז בתוך אפליקציית הדסקטופ — הגרסה החדשה ביותר
            string bundled = Path.Combine(appData, @"Claude\claude-code");
            if (Directory.Exists(bundled))
            {
                var dirs = new List<string>(Directory.GetDirectories(bundled));
                dirs.Sort(StringComparer.OrdinalIgnoreCase);
                for (int i = dirs.Count - 1; i >= 0; i--)
                {
                    string exe = Path.Combine(dirs[i], "claude.exe");
                    if (File.Exists(exe)) yield return exe;
                }
            }
        }

        public static string Locate()
        {
            foreach (string p in CandidatePaths())
                if (!string.IsNullOrEmpty(p) && File.Exists(p)) return p;

            string fromPath = SearchPath("claude.cmd") ?? SearchPath("claude.exe");
            return fromPath;
        }

        private static string SearchPath(string fileName)
        {
            string path = Environment.GetEnvironmentVariable("PATH");
            if (string.IsNullOrEmpty(path)) return null;
            foreach (string dir in path.Split(';'))
            {
                if (string.IsNullOrEmpty(dir)) continue;
                try
                {
                    string full = Path.Combine(dir.Trim(), fileName);
                    if (File.Exists(full)) return full;
                }
                catch { /* רכיב נתיב לא חוקי */ }
            }
            return null;
        }

        // ---------- זמינות ----------

        public string CheckAvailability()
        {
            if (string.IsNullOrEmpty(_exePath))
                return "לא נמצאה התקנה של Claude Code במחשב. פתחו את אשף החיבור בתוסף כדי להתקין ולהתחבר.";

            if (!File.Exists(_exePath))
                return "הנתיב ל-Claude Code שמור אך הקובץ אינו קיים: " + _exePath;

            try
            {
                string outText, errText;
                int exit = Run(new[] { "--version" }, null, 20000, out outText, out errText);
                if (exit != 0)
                    return "Claude Code נמצא אך אינו מגיב כראוי. פרטים: " + Trim(errText, 200);
            }
            catch (Exception ex)
            {
                return "לא ניתן להריץ את Claude Code: " + ex.Message;
            }

            // בדיקת אימות בפועל — הרצה מינימלית שמחזירה תשובה קצרה
            try
            {
                string outText, errText;
                int exit = Run(
                    new[] { "-p", "--output-format", "json", "--max-turns", "1", "--allowed-tools", "" },
                    "החזר בדיוק את המילה: תקין",
                    60000, out outText, out errText);

                string envelope = ReadEnvelopeResult(outText);
                if (envelope != null && envelope.IndexOf("Not logged in", StringComparison.OrdinalIgnoreCase) >= 0)
                    return "Claude Code מותקן אך אינו מחובר לחשבון. הריצו בחלון פקודה: claude  ואז /login";

                if (exit != 0 && envelope == null)
                    return "בדיקת החיבור ל-Claude Code נכשלה. פרטים: " + Trim(errText ?? outText, 200);
            }
            catch (Exception ex)
            {
                return "בדיקת החיבור ל-Claude Code נכשלה: " + ex.Message;
            }

            return null;
        }

        // ---------- הרצה ----------

        public string Decide(string skeletonJson, RulesFile rules)
        {
            if (string.IsNullOrEmpty(_exePath))
                throw new EngineException("Claude Code אינו מותקן או שלא אותר במחשב.");

            // הנחיית המערכת עוברת בקובץ ולא בשורת הפקודה.
            // היא מכילה גרשיים ותווי JSON, ושורת פקודה היא המקום הכי שביר להעביר בו כאלה.
            string promptFile = Path.Combine(Path.GetTempPath(),
                "sadar-sys-" + Guid.NewGuid().ToString("N") + ".txt");

            string outText, errText;
            int exit;

            try
            {
                File.WriteAllText(promptFile, SystemPrompt, new UTF8Encoding(false));

                var args = new List<string>
                {
                    "-p",
                    "--output-format", "json",
                    "--max-turns", "1",
                    "--allowed-tools", "",
                    "--model", _model,
                    "--system-prompt-file", promptFile
                };

                exit = Run(args.ToArray(), skeletonJson, _timeoutMs, out outText, out errText);
            }
            finally
            {
                try { if (File.Exists(promptFile)) File.Delete(promptFile); } catch { }
            }

            string result = ReadEnvelopeResult(outText);

            if (result == null)
            {
                if (LooksLikeFilter(errText) || LooksLikeFilter(outText))
                    throw new EngineException(
                        "החיבור נחסם על ידי הסינון ברשת. נסו להפעיל 'מצב מסונן' בהגדרות התוסף.",
                        true);

                throw new EngineException(
                    "לא התקבלה תשובה תקינה מ-Claude Code. קוד יציאה " + exit + ". " +
                    Trim(errText ?? outText, 300));
            }

            if (result.IndexOf("Not logged in", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new EngineException(
                    "Claude Code אינו מחובר לחשבון. הריצו בחלון פקודה: claude  ואז /login");

            string json = Json.Json.ExtractObject(result);
            if (json == null)
                throw new EngineException("התשובה מ-Claude Code אינה מכילה JSON תקין.");

            return json;
        }

        /// <summary>שולף את שדה result מהמעטפת שה-CLI מחזיר ב---output-format json.</summary>
        private static string ReadEnvelopeResult(string stdout)
        {
            if (string.IsNullOrWhiteSpace(stdout)) return null;
            try
            {
                var env = Json.Json.ParseObject(stdout.Trim());
                return Json.Json.GetString(env, "result", null);
            }
            catch
            {
                return null;
            }
        }

        private int Run(string[] args, string stdin, int timeoutMs, out string stdout, out string stderr)
        {
            bool isCmd = _exePath.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase) ||
                         _exePath.EndsWith(".bat", StringComparison.OrdinalIgnoreCase);

            var psi = new ProcessStartInfo
            {
                FileName = isCmd ? "cmd.exe" : _exePath,
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
            if (isCmd) sb.Append("/c ").Append(Quote(_exePath));
            foreach (string a in args)
            {
                sb.Append(' ');
                sb.Append(Quote(a));
            }
            psi.Arguments = sb.ToString().TrimStart();

            // הסוכן לא אמור לרשת שום הקשר מהתהליך שקרא לו
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
                    // תהליך שנפל מוקדם סוגר את הצינור, והכתיבה זורקת.
                    // זו אינה השגיאה האמיתית — היא רק מסתירה את הסיבה,
                    // שתופיע ב-stderr ותיקרא בהמשך.
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
                        "Claude Code לא השיב בתוך " + (timeoutMs / 1000) + " שניות. החלון בוטל והמסמך לא נגע.");
                }

                proc.WaitForExit(); // השלמת קריאת הפלט האסינכרוני

                stdout = outBuf.ToString();
                stderr = errBuf.ToString();
                return proc.ExitCode;
            }
        }

        /// <summary>
        /// ציטוט לפי כללי שורת הפקודה של Windows.
        ///
        /// רק לוכסן אחורי שקודם לגרש דורש הכפלה. הכפלת כל הלוכסנים —
        /// טעות נפוצה — הורסת כל נתיב קובץ שמועבר כארגומנט.
        /// </summary>
        private static string Quote(string s)
        {
            if (string.IsNullOrEmpty(s)) return "\"\"";

            var sb = new StringBuilder(s.Length + 8);
            sb.Append('"');

            int backslashes = 0;
            foreach (char c in s)
            {
                if (c == '\\') { backslashes++; continue; }

                if (c == '"')
                {
                    sb.Append('\\', backslashes * 2 + 1).Append('"');
                }
                else
                {
                    sb.Append('\\', backslashes);
                    sb.Append(c);
                }
                backslashes = 0;
            }

            // לוכסנים בסוף מוכפלים, אחרת הם יבריחו את הגרש הסוגר
            sb.Append('\\', backslashes * 2);
            sb.Append('"');
            return sb.ToString();
        }

        /// <summary>
        /// זיהוי חסימת סינון מול תקלת רשת. ההבחנה חשובה:
        /// "השרת נפל" ו"נטפרי חסמה" דורשים פעולה אחרת לגמרי מהמשתמש.
        /// </summary>
        public static bool LooksLikeFilter(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            string t = text.ToLowerInvariant();
            return t.Contains("netfree") ||
                   t.Contains("rimon") ||
                   t.Contains("self signed certificate") ||
                   t.Contains("self-signed certificate") ||
                   t.Contains("unable to verify the first certificate") ||
                   t.Contains("cert_authority_invalid") ||
                   t.Contains("blocked") ||
                   text.Contains("חסום") ||
                   text.Contains("חסימה");
        }

        private static string Trim(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Trim();
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }
    }
}
