using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Sadar.Core;
using Sadar.Core.Document;
using Sadar.Core.Engine;
using Sadar.Core.Model;
using Sadar.Core.Rules;

namespace Sadar.Cli
{
    /// <summary>
    /// כלי שורת פקודה. שני תפקידים:
    ///   1. מערך הבדיקות של הליבה — רץ בלי וורד ובלי רשת.
    ///   2. עיבוד אצווה של קבצי docx למי שמעדיף כך.
    /// </summary>
    public static class Program
    {
        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.UTF8;

            if (args.Length == 0) { Usage(); return 1; }

            try
            {
                switch (args[0].ToLowerInvariant())
                {
                    case "sample": return CmdSample(args);
                    case "scan": return CmdScan(args);
                    case "analyze": return CmdAnalyze(args, false);
                    case "apply": return CmdAnalyze(args, true);
                    case "clean": return CmdClean(args);
                    case "rules": return CmdRules(args);
                    case "doctor": return CmdDoctor();
                    case "engines": return CmdEngines(args);
                    case "key": return CmdKey(args);
                    case "selftest": return SelfTest.Run();
                    case "qa": return QaTests.Run();
                    default: Usage(); return 1;
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("שגיאה: " + ex.Message);
                Console.ResetColor();
                return 2;
            }
        }

        private static void Usage()
        {
            Console.WriteLine(@"
סַדָּר — עימוד חכם למסמכי וורד

  sadar sample <out.docx>              יוצר מסמך בדיקה
  sadar scan <file.docx>               מציג את השלד שיישלח למנוע
  sadar analyze <file.docx> [אפשרויות] מנתח ומציג הצעה, בלי לשנות דבר
  sadar apply <file.docx> [אפשרויות]   מנתח, מחיל ומאמת
  sadar clean <file.docx> [--out X]    ניקוי דטרמיניסטי בלבד, בלי מנוע
  sadar rules init <out.json>          יוצר קובץ כללים לדוגמה
  sadar engines                        רשימת המנועים ומצב כל אחד
  sadar engines use <id>               בוחר מנוע ברירת מחדל
  sadar key <id> <מפתח>                שומר מפתח API (מוצפן, למשתמש הזה בלבד)
  sadar key <id> --clear               מוחק מפתח
  sadar doctor                         בודק את כל המנועים
  sadar selftest                       מריץ את בדיקות הליבה
  sadar qa                             בדיקות קצה והתנהגות עוינת

אפשרויות:
  --engine <id>          מנוע לשימוש. mock = בדיקות בלי רשת.
                         ברירת המחדל נלקחת מההגדרות השמורות
  --rules <file.json>    קובץ כללים
  --out <file.docx>      קובץ יעד (ברירת מחדל: דריסה עם גיבוי)
  --model <name>         מודל, ברירת מחדל sonnet
  --privacy              מצב מסונן: לא נשלח טקסט כלל
");
        }

        // ---------- פקודות ----------

        private static int CmdSample(string[] args)
        {
            string path = Arg(args, 1) ?? "sample.docx";
            int repeat = 1;
            string r = Opt(args, "--repeat");
            if (r != null) int.TryParse(r, out repeat);
            SampleDocument.Create(path, Math.Max(1, repeat));
            Console.WriteLine("נוצר מסמך בדיקה: " + path);
            return 0;
        }

        private static int CmdScan(string[] args)
        {
            string path = Require(args, 1, "נתיב קובץ docx");
            var rules = LoadRules(args);

            using (var doc = new OpenXmlDocumentAdapter(path))
            {
                var paragraphs = doc.ReadAll();
                string skeleton = Core.Scanner.SkeletonBuilder.Build(
                    paragraphs, 0, paragraphs.Count, rules, null);

                Console.WriteLine(skeleton);
                Console.Error.WriteLine();
                Console.Error.WriteLine("פסקאות: " + paragraphs.Count);
                Console.Error.WriteLine("נפח השלד: " + skeleton.Length + " תווים");

                int fullText = 0;
                foreach (var p in paragraphs) fullText += p.Length;
                Console.Error.WriteLine("נפח הטקסט המלא: " + fullText + " תווים");
                if (fullText > 0)
                    Console.Error.WriteLine("נשלח למנוע: " +
                        Math.Round(100.0 * skeleton.Length / fullText, 1) + "% מנפח המסמך");
            }
            return 0;
        }

        private static int CmdAnalyze(string[] args, bool apply)
        {
            string path = Require(args, 1, "נתיב קובץ docx");
            var rules = LoadRules(args);
            string engineName = Opt(args, "--engine");
            string outPath = Opt(args, "--out");

            using (var doc = new OpenXmlDocumentAdapter(path))
            {
                var snapshot = doc.ReadAll();

                var engine = BuildEngine(engineName, args, snapshot);

                Console.WriteLine("מנוע: " + engine.Name);

                var orch = new Orchestrator(doc, engine, rules);
                orch.Progress += (cur, total, msg) => Console.WriteLine("  " + msg);

                var proposal = orch.Analyze(true);
                PrintProposal(proposal, rules);

                if (!apply) return 0;

                if (outPath == null)
                {
                    string backup = MakeBackupPath(path);
                    File.Copy(path, backup, true);
                    Console.WriteLine();
                    Console.WriteLine("גיבוי נשמר: " + Path.GetFileName(backup));
                }

                var result = orch.Apply(proposal, null, true);

                Console.WriteLine();
                if (result.Success)
                {
                    doc.SaveAs(outPath ?? path);
                    PrintReport(result.Report, outPath ?? path);
                    return 0;
                }

                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ההחלה בוטלה: " + result.ErrorMessage);
                Console.ResetColor();
                Console.WriteLine("המסמך לא נשמר ולא השתנה.");
                return 3;
            }
        }

        private static int CmdClean(string[] args)
        {
            string path = Require(args, 1, "נתיב קובץ docx");
            var rules = LoadRules(args);
            string outPath = Opt(args, "--out");

            using (var doc = new OpenXmlDocumentAdapter(path))
            {
                var orch = new Orchestrator(doc, null, rules);
                var proposal = orch.Analyze(false);

                Console.WriteLine("ניקוי דטרמיניסטי (בלי מנוע, בלי רשת):");
                Console.WriteLine("  תיקוני רווחים: " + proposal.CleanPlan.Rewrites.Count);
                Console.WriteLine("  פסקאות ריקות למחיקה: " + proposal.CleanPlan.Deletes.Count);

                if (proposal.CleanPlan.TotalChanges == 0)
                {
                    Console.WriteLine("אין מה לנקות.");
                    return 0;
                }

                if (outPath == null)
                {
                    string backup = MakeBackupPath(path);
                    File.Copy(path, backup, true);
                    Console.WriteLine("גיבוי נשמר: " + Path.GetFileName(backup));
                }

                var result = orch.Apply(proposal, new HashSet<int>(), true);
                if (!result.Success)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("בוטל: " + result.ErrorMessage);
                    Console.ResetColor();
                    return 3;
                }

                doc.SaveAs(outPath ?? path);
                PrintReport(result.Report, outPath ?? path);
                return 0;
            }
        }

        private static int CmdRules(string[] args)
        {
            string sub = Arg(args, 1);
            if (!string.Equals(sub, "init", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("שימוש: sadar rules init <out.json>");
                return 1;
            }

            string path = Arg(args, 2) ?? "כללים.json";
            var rules = new RulesFile
            {
                Name = "ספר שאלות ותשובות",
                Description = "כללים לדוגמה — ערכו לפי הסדרה שלכם",
                StructureHint =
                    "הספר מחולק לסימנים. פסקה קצרה ומודגשת שמתחילה במילה 'סימן' היא כותרת ראשית. " +
                    "הפסקה המודגשת שאחריה, שהיא תיאור נושא הסימן, היא כותרת משנה. " +
                    "כל השאר הוא גוף הטקסט."
            };
            rules.NeverTouchPrefixes.Add("בס\"ד");
            rules.Save(path);

            Console.WriteLine("נוצר קובץ כללים: " + path);
            return 0;
        }

        /// <summary>
        /// בוחר מנוע: הדגל מנצח, אחריו ההגדרות השמורות.
        /// "mock" נשאר תמיד זמין — הוא מה שמאפשר לבדוק את כל הצינור בלי רשת.
        /// </summary>
        private static IDecisionEngine BuildEngine(string engineName, string[] args,
                                                   System.Collections.Generic.List<ParagraphInfo> snapshot)
        {
            if (string.Equals(engineName, "mock", StringComparison.OrdinalIgnoreCase))
                return new MockEngine(snapshot);

            var settings = EngineSettings.LoadDefault();

            string model = Opt(args, "--model");
            if (!string.IsNullOrEmpty(model)) settings.Model = model;

            string timeout = Opt(args, "--timeout");
            int seconds;
            if (timeout != null && int.TryParse(timeout, out seconds)) settings.TimeoutSeconds = seconds;

            if (string.IsNullOrEmpty(engineName)) return EngineFactory.Create(settings);

            // "claude" נשמר כשם נוח למנוע ברירת המחדל
            if (string.Equals(engineName, "claude", StringComparison.OrdinalIgnoreCase))
                engineName = "claude-cli";

            return EngineFactory.CreateById(engineName, settings);
        }

        private static int CmdEngines(string[] args)
        {
            var settings = EngineSettings.LoadDefault();

            // בחירת מנוע ברירת מחדל
            if (string.Equals(Arg(args, 1), "use", StringComparison.OrdinalIgnoreCase))
            {
                string id = Arg(args, 2);
                var info = EngineInfo.ById(id);
                if (info == null)
                {
                    Console.WriteLine("מנוע לא מוכר: " + (id ?? "(חסר)"));
                    Console.WriteLine("הרשימה: sadar engines");
                    return 1;
                }

                settings.EngineId = info.Id;
                settings.SaveDefault();
                Console.WriteLine("המנוע נקבע ל: " + info.DisplayName);
                if (info.NeedsApiKey && !settings.HasKey(info.Id))
                    Console.WriteLine("עדיין חסר מפתח. הריצו: sadar key " + info.Id + " <מפתח>");
                return 0;
            }

            bool deep = Has(args, "--check");

            Console.WriteLine("המנוע הנוכחי: " + settings.Info.DisplayName +
                              "   (מודל: " + settings.EffectiveModel + ")");
            Console.WriteLine();
            Console.WriteLine(deep
                ? "בודק את כל המנועים, זה עשוי לקחת רגע..."
                : "להוספת בדיקת חיבור בפועל: sadar engines --check");
            Console.WriteLine();

            foreach (var st in EngineFactory.Survey(settings, deep))
            {
                bool current = st.Info.Id == settings.EngineId;
                Console.Write(current ? " * " : "   ");
                Console.Write(st.Info.Id.PadRight(16));
                Console.Write(st.Info.DisplayName.PadRight(24));

                if (st.Problem == null && st.Configured)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(deep ? "מוכן" : "מוגדר");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write(st.Problem ?? "לא מוגדר");
                }
                Console.ResetColor();
                Console.WriteLine();

                if (!string.IsNullOrEmpty(st.Detail))
                    Console.WriteLine("        " + st.Detail);
                else if (!st.Configured)
                    Console.WriteLine("        " + st.Info.HowToConnect);
            }

            Console.WriteLine();
            Console.WriteLine("ההגדרות נשמרות ב: " + EngineSettings.DefaultPath);
            return 0;
        }

        private static int CmdKey(string[] args)
        {
            string id = Arg(args, 1);
            var info = EngineInfo.ById(id);
            if (info == null)
            {
                Console.WriteLine("שימוש: sadar key <מזהה מנוע> <מפתח>");
                Console.WriteLine("הרשימה: sadar engines");
                return 1;
            }

            if (!info.NeedsApiKey)
            {
                Console.WriteLine(info.DisplayName + " עובד דרך המנוי ואינו צריך מפתח.");
                Console.WriteLine(info.HowToConnect);
                return 1;
            }

            var settings = EngineSettings.LoadDefault();

            if (Has(args, "--clear"))
            {
                settings.SetKey(info.Id, null);
                settings.SaveDefault();
                Console.WriteLine("המפתch נמחק.".Replace("ch", "ח"));
                return 0;
            }

            string key = Arg(args, 2);
            if (string.IsNullOrEmpty(key))
            {
                Console.WriteLine("שימוש: sadar key " + info.Id + " <מפתח>");
                return 1;
            }

            settings.SetKey(info.Id, key);
            if (!settings.HasKey(info.Id))
            {
                Console.WriteLine("שמירת המפתח נכשלה — הצפנת Windows אינה זמינה.");
                return 2;
            }

            settings.SaveDefault();
            Console.WriteLine("המפתח נשמר עבור " + info.DisplayName + " " + settings.KeyHint(info.Id));
            Console.WriteLine("הוא מוצפן למשתמש הזה בלבד ואינו קריא ממחשב אחר.");
            return 0;
        }

        private static int CmdDoctor()
        {
            var settings = EngineSettings.LoadDefault();

            Console.WriteLine("בודק את הסביבה...");
            Console.WriteLine();
            Console.WriteLine("המנוע שנבחר: " + settings.Info.DisplayName);
            Console.WriteLine("מודל: " + settings.EffectiveModel);
            Console.WriteLine();

            int ready = 0;
            foreach (var st in EngineFactory.Survey(settings, true))
            {
                bool current = st.Info.Id == settings.EngineId;
                if (st.Ready && st.Configured) ready++;

                Console.Write(current ? " * " : "   ");
                Console.Write(st.Info.DisplayName.PadRight(26));

                if (st.Ready && st.Configured)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("מוכן");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(st.Problem ?? "לא מוגדר");
                }
                Console.ResetColor();

                if (!string.IsNullOrEmpty(st.Detail)) Console.WriteLine("      " + st.Detail);
            }

            Console.WriteLine();
            if (ready == 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("אף מנוע אינו מוכן. \"ניקוי בלבד\" עדיין עובד — הוא רץ מקומית.");
                Console.ResetColor();
                return 1;
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(ready + " מנועים מוכנים לשימוש.");
            Console.ResetColor();
            return 0;
        }

        /// <summary>
        /// בוחר מנוע: הדגל מנצח, אחריו ההגדרות השמורות.
        /// "mock" נשאר תמיד זמין — הוא מה שמאפשר לבדוק את כל הצינור בלי רשת.
        /// </summary>
        // ---------- תצוגה ----------

        private static void PrintProposal(Proposal proposal, RulesFile rules)
        {
            Console.WriteLine();
            Console.WriteLine("---- הצעה ----");
            Console.WriteLine("פסקאות במסמך: " + proposal.Snapshot.Count);
            Console.WriteLine("חלונות שעובדו: " + proposal.WindowsProcessed);
            Console.WriteLine("זמן ניתוח: " + Math.Round(proposal.AnalyzeDuration.TotalSeconds, 1) + " שניות");
            Console.WriteLine();

            var changes = proposal.EffectiveChanges(rules);
            Console.WriteLine("שינויי סגנון: " + changes.Count);

            int shown = 0;
            foreach (var d in changes)
            {
                if (shown++ >= 25) { Console.WriteLine("  ... ועוד " + (changes.Count - 25)); break; }

                var p = FindParagraph(proposal.Snapshot, d.Index);
                string preview = p == null ? "" : p.Prefix(45);
                Console.WriteLine(string.Format("  פסקה {0,4}: {1,-9} -> {2,-9} | {3}{4}",
                    d.Index + 1,
                    p == null ? "?" : p.StyleName,
                    d.StyleName,
                    preview,
                    d.IsUncertain ? "   [לבדיקה: " + d.Reason + "]" : ""));
            }

            Console.WriteLine();
            Console.WriteLine("ניקוי: " + proposal.CleanPlan.Rewrites.Count + " תיקוני רווחים, " +
                              proposal.CleanPlan.Deletes.Count + " פסקאות ריקות למחיקה");

            if (proposal.Warnings.Count > 0)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("הערות:");
                foreach (string w in proposal.Warnings) Console.WriteLine("  - " + w);
                Console.ResetColor();
            }
        }

        private static void PrintReport(RunReport r, string savedTo)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("הושלם.");
            Console.ResetColor();
            Console.WriteLine("  סגנונות שהוחלו: " + r.StylesApplied);
            foreach (var kv in r.StyleCounts)
                Console.WriteLine("      " + kv.Key + ": " + kv.Value);
            Console.WriteLine("  תיקוני רווחים: " + r.WhitespaceFixes);
            Console.WriteLine("  פסקאות שנמחקו: " + r.ParagraphsDeleted);
            Console.WriteLine("  סומנו לבדיקה: " + r.UncertainCount);
            Console.WriteLine("  אימות המלל: " + (r.TextVerified ? "עבר — המלל לא השתנה" : "נכשל"));
            Console.WriteLine("  נשמר: " + savedTo);
        }

        private static ParagraphInfo FindParagraph(List<ParagraphInfo> list, int index)
        {
            foreach (var p in list) if (p.Index == index) return p;
            return null;
        }

        // ---------- עזר ----------

        private static string MakeBackupPath(string path)
        {
            string dir = Path.GetDirectoryName(Path.GetFullPath(path));
            string name = Path.GetFileNameWithoutExtension(path);
            string stamp = DateTime.Now.ToString("yyyy-MM-dd_HHmm");
            return Path.Combine(dir, name + ".גיבוי-לפני-סידור-" + stamp + ".docx");
        }

        private static RulesFile LoadRules(string[] args)
        {
            string path = Opt(args, "--rules");
            var rules = path != null ? RulesFile.Load(path) : new RulesFile();
            if (Has(args, "--privacy")) rules.PrivacyMode = true;
            return rules;
        }

        private static string Arg(string[] args, int i)
        {
            return i < args.Length && !args[i].StartsWith("--") ? args[i] : null;
        }

        private static string Require(string[] args, int i, string what)
        {
            string v = Arg(args, i);
            if (v == null) throw new ArgumentException("חסר: " + what);
            return v;
        }

        private static string Opt(string[] args, string name)
        {
            for (int i = 0; i < args.Length - 1; i++)
                if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
                    return args[i + 1];
            return null;
        }

        private static bool Has(string[] args, string name)
        {
            foreach (string a in args)
                if (string.Equals(a, name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }
    }
}
