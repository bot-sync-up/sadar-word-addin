using System;
using System.Collections.Generic;
using System.IO;
using Sadar.Core;
using Sadar.Core.Document;
using Sadar.Core.Engine;
using Sadar.Core.Model;
using Sadar.Core.Rules;
using Sadar.Core.Safety;
using Sadar.Core.Windowing;

namespace Sadar.Cli
{
    /// <summary>
    /// בדיקות הליבה. הדגש הוא על חוזה הבטיחות:
    /// רוב הבדיקות כאן מוודאות שהמערכת מסרבת לעשות דברים, לא שהיא עושה אותם.
    /// </summary>
    public static class SelfTest
    {
        private static int _passed, _failed;

        public static int Run()
        {
            Console.WriteLine("בדיקות סַדָּר");
            Console.WriteLine(new string('=', 60));

            TestNormalization();
            TestHashSensitivity();
            TestCleaner();
            TestParserRejectsBadStyle();
            TestParserRejectsDeleteOfNonEmpty();
            TestParserRejectsOutOfRange();
            TestProtectedParagraphs();
            TestWindowPlanner();
            TestWindowMergeConflict();
            TestPrivacyFingerprint();
            TestEndToEnd();
            TestVerifierCatchesTextChange();

            Console.WriteLine(new string('=', 60));
            Console.ForegroundColor = _failed == 0 ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine(string.Format("עברו {0}, נכשלו {1}", _passed, _failed));
            Console.ResetColor();
            return _failed == 0 ? 0 : 1;
        }

        // ---------- בדיקות ----------

        private static void TestNormalization()
        {
            Check("נרמול מכווץ רווחים כפולים",
                TextHasher.Normalize("שלום   עולם") == "שלום עולם");

            Check("נרמול חותך רווחים בקצוות",
                TextHasher.Normalize("  שלום  ") == "שלום");

            Check("נרמול מסיר תווי בקרה של וורד",
                TextHasher.Normalize("שלום\r\a") == "שלום");

            Check("נרמול שומר על גרשיים",
                TextHasher.Normalize("הרמב\"ם") == "הרמב\"ם");
        }

        private static void TestHashSensitivity()
        {
            var a = Paras("שלום עולם", "", "טקסט שני");
            var b = Paras("שלום    עולם", "טקסט   שני");
            Check("hash אדיש לרווחים ולפסקאות ריקות",
                TextHasher.HashDocument(a) == TextHasher.HashDocument(b));

            var c = Paras("שלום עולם", "טקסט שני!");
            Check("hash רגיש לשינוי פיסוק",
                TextHasher.HashDocument(a) != TextHasher.HashDocument(c));

            var d = Paras("שלום עולם", "טקסט שנים");
            Check("hash רגיש לשינוי אות אחת",
                TextHasher.HashDocument(a) != TextHasher.HashDocument(d));

            var e = Paras("שלום עולם", "טקסט שֵני");
            Check("hash רגיש לתוספת ניקוד",
                TextHasher.HashDocument(a) != TextHasher.HashDocument(e));
        }

        private static void TestCleaner()
        {
            var rules = new RulesFile();

            Check("ניקוי מכווץ רווחים",
                DeterministicCleaner.CleanText("א  ב   ג", rules) == "א ב ג");

            Check("ניקוי מסיר רווח לפני נקודה",
                DeterministicCleaner.CleanText("סוף המשפט .", rules) == "סוף המשפט.");

            Check("ניקוי אינו נוגע בגרשיים של ראשי תיבות",
                DeterministicCleaner.CleanText("שו\"ע או\"ח", rules) == "שו\"ע או\"ח");

            Check("ניקוי הופך טאב לרווח",
                DeterministicCleaner.CleanText("א\tב", rules) == "א ב");

            // הבדיקה החשובה: הניקוי לא משנה את טביעת האצבע
            string original = "  הרמב\"ם   כתב  כך .  ";
            string cleaned = DeterministicCleaner.CleanText(original, rules);
            Check("ניקוי אינו משנה את טביעת האצבע של המלל",
                TextHasher.Normalize(original) == TextHasher.Normalize(cleaned + " "));
        }

        private static void TestParserRejectsBadStyle()
        {
            var rules = new RulesFile();
            var allowed = Allowed(Paras("כותרת כלשהי"));

            var set = DecisionParser.Parse(
                "{\"decisions\":[{\"i\":0,\"style\":\"סגנון שלא קיים\"}]}", allowed, rules);

            Check("נפסל סגנון שאינו ברשימה המותרת",
                set.Decisions.Count == 0 && set.Warnings.Count == 1);
        }

        private static void TestParserRejectsDeleteOfNonEmpty()
        {
            var rules = new RulesFile();
            var allowed = Allowed(Paras("פסקה עם טקסט אמיתי"));

            var set = DecisionParser.Parse("{\"deletes\":[0]}", allowed, rules);

            Check("נפסלה מחיקה של פסקה שיש בה טקסט",
                set.Deletes.Count == 0 && set.Warnings.Count == 1);
        }

        private static void TestParserRejectsOutOfRange()
        {
            var rules = new RulesFile();
            var allowed = Allowed(Paras("פסקה"));

            var set = DecisionParser.Parse(
                "{\"decisions\":[{\"i\":999,\"style\":\"כותרת 1\"}]}", allowed, rules);

            Check("נפסלה החלטה על פסקה מחוץ לחלון",
                set.Decisions.Count == 0);
        }

        private static void TestProtectedParagraphs()
        {
            var rules = new RulesFile();
            rules.NeverTouchPrefixes.Add("בס\"ד");

            var paras = Paras("בס\"ד ראש הדף");
            var allowed = Allowed(paras);

            var set = DecisionParser.Parse(
                "{\"decisions\":[{\"i\":0,\"style\":\"כותרת 1\"}]}", allowed, rules);

            Check("פסקה מוגנת לא קיבלה סגנון",
                set.Decisions.Count == 0);
        }

        private static void TestWindowPlanner()
        {
            var rules = new RulesFile { WindowSize = 100, WindowOverlap = 20 };

            var one = WindowPlanner.Plan(50, rules);
            Check("מסמך קטן = חלון אחד", one.Count == 1 && one[0].Count == 50);

            var many = WindowPlanner.Plan(500, rules);
            bool coversAll = many[many.Count - 1].End >= 500;
            bool hasOverlap = many.Count > 1 && many[1].Start < many[0].End;
            Check("חלוקה לחלונות מכסה את כל המסמך", coversAll);
            Check("חלונות עוקבים חופפים", hasOverlap);
        }

        private static void TestWindowMergeConflict()
        {
            var a = new DecisionSet();
            a.Decisions.Add(new StyleDecision { Index = 5, StyleName = "כותרת 1" });

            var b = new DecisionSet();
            b.Decisions.Add(new StyleDecision { Index = 5, StyleName = "כותרת 2" });

            var merged = DecisionMerger.Merge(new[] { a, b });

            Check("אי-הסכמה בין חלונות מסומנת לבדיקה ולא מוכרעת שרירותית",
                merged.Decisions.Count == 1 && merged.Decisions[0].IsUncertain);
        }

        private static void TestPrivacyFingerprint()
        {
            string fp = Core.Scanner.SkeletonBuilder.Fingerprint("סימן א");

            bool leaksText = fp.Contains("סימן") || fp.Contains("א");
            Check("טביעת אצבע במצב מסונן אינה מכילה טקסט", !leaksText);
            Check("טביעת אצבע מקודדת מספר מילים", fp.StartsWith("W2"));
        }

        private static void TestEndToEnd()
        {
            string dir = Path.Combine(Path.GetTempPath(), "sadar-selftest");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "e2e.docx");

            try
            {
                SampleDocument.Create(path);

                using (var doc = new OpenXmlDocumentAdapter(path))
                {
                    var before = doc.ReadAll();
                    string hashBefore = TextHasher.HashDocument(before);

                    var rules = new RulesFile();
                    rules.NeverTouchPrefixes.Add("בס\"ד");

                    var orch = new Orchestrator(doc, new MockEngine(before), rules);
                    var proposal = orch.Analyze(true);

                    Check("הניתוח זיהה כותרות", proposal.EffectiveChanges(rules).Count >= 6);
                    Check("הניתוח מצא רווחים כפולים לניקוי", proposal.CleanPlan.Rewrites.Count > 0);
                    Check("הניתוח מצא פסקאות ריקות עודפות", proposal.CleanPlan.Deletes.Count > 0);

                    var result = orch.Apply(proposal, null, true);

                    Check("ההחלה הצליחה", result.Success);
                    Check("אימות המלל עבר", result.Report.TextVerified);

                    var after = doc.ReadAll();
                    Check("המלל לא השתנה אחרי ההחלה",
                        TextHasher.HashDocument(after) == hashBefore);

                    int headings = 0;
                    foreach (var p in after)
                        if (p.StyleName != null && p.StyleName.StartsWith("כותרת")) headings++;
                    Check("סגנונות כותרת הוחלו בפועל", headings >= 6);

                    string outPath = Path.Combine(dir, "e2e-out.docx");
                    doc.SaveAs(outPath);

                    using (var reopened = new OpenXmlDocumentAdapter(outPath))
                    {
                        var reread = reopened.ReadAll();
                        Check("המלל שרד סבב שמירה וטעינה",
                            TextHasher.HashDocument(reread) == hashBefore);

                        int h2 = 0;
                        foreach (var p in reread)
                            if (p.StyleName != null && p.StyleName.StartsWith("כותרת")) h2++;
                        Check("הסגנונות שרדו סבב שמירה וטעינה", h2 >= 6);
                    }
                }
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        private static void TestVerifierCatchesTextChange()
        {
            var before = Paras("שלום עולם", "טקסט שני");
            var after = Paras("שלום עולם", "טקסט שלישי");

            var v = TextVerifier.Verify(before, after);

            Check("המאמת תופס שינוי במלל", !v.Passed);
            Check("המאמת מצביע על הפסקה הנכונה", v.FirstChangedParagraph == 1);
        }

        // ---------- עזר ----------

        private static List<ParagraphInfo> Paras(params string[] texts)
        {
            var list = new List<ParagraphInfo>();
            for (int i = 0; i < texts.Length; i++)
                list.Add(new ParagraphInfo
                {
                    Index = i,
                    Text = texts[i],
                    StyleName = "רגיל",
                    FontSize = 12
                });
            return list;
        }

        private static Dictionary<int, ParagraphInfo> Allowed(List<ParagraphInfo> ps)
        {
            var d = new Dictionary<int, ParagraphInfo>();
            foreach (var p in ps) d[p.Index] = p;
            return d;
        }

        private static void Check(string name, bool condition)
        {
            if (condition)
            {
                _passed++;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("  עבר   ");
            }
            else
            {
                _failed++;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("  נכשל  ");
            }
            Console.ResetColor();
            Console.WriteLine(name);
        }
    }
}
