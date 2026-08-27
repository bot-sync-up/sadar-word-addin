using System;
using System.Collections.Generic;
using System.IO;
using Sadar.Core;
using Sadar.Core.Document;
using Sadar.Core.Engine;
using Sadar.Core.Model;
using Sadar.Core.Rules;
using Sadar.Core.Safety;
using Sadar.Core.Scanner;
using Sadar.Core.Windowing;

namespace Sadar.Cli
{
    /// <summary>
    /// בדיקות קצה ועוינות.
    ///
    /// selftest בודק שהמערכת עושה את מה שהיא אמורה. כאן נבדק ההפך:
    /// מה קורה במסמכים מנוונים, ומה קורה כשמנוע ההחלטות מתנהג רע
    /// או כשמשהו בצינור מנסה לגעת במלל.
    /// </summary>
    public static class QaTests
    {
        private static int _passed, _failed;
        private static readonly List<string> Failures = new List<string>();

        public static int Run()
        {
            _passed = 0; _failed = 0; Failures.Clear();

            Console.WriteLine("בדיקות QA — מקרי קצה והתנהגות עוינת");
            Console.WriteLine(new string('=', 66));

            Section("מסמכים מנוונים");
            TestEmptyDocument();
            TestSingleParagraph();
            TestAllEmptyParagraphs();
            TestNoEmptyParagraphs();

            Section("שלד");
            TestSkeletonIsValidJson();
            TestLongParagraphTextNotSent();
            TestPrivacyModeLeaksNothing();
            TestSkeletonHandlesEmptyWindow();

            Section("מנוע עוין");
            TestMalformedJson();
            TestNotJsonAtAll();
            TestEmptyResponse();
            TestOutOfRangeIndex();
            TestDuplicateDecisions();
            TestDeleteEverything();
            TestUnknownStyle();
            TestJsonWrappedInMarkdown();

            Section("חוזה הבטיחות");
            TestSabotageIsCaughtAndRolledBack();
            TestNikudChangeIsCaught();
            TestGershayimSurvive();
            TestCleanerIsIdempotent();

            Section("עיצוב פנימי בפסקה");
            TestInlineBoldSurvivesCleaning();
            TestSpaceRemovedAcrossRunBoundary();
            TestTabsAndNbspCleaned();

            QaEngineTests.Run(Section, Check);

            Section("חלונות");
            TestWindowsCoverExactMultiple();
            TestWindowsCoverOffByOne();
            TestAgreementUpgradesCertainty();

            Section("כללים");
            TestRulesRoundTrip();
            TestRulesRejectsGarbage();
            TestRulesClampsInsaneValues();

            Section("החלה");
            TestApplyOnlyCleaning();
            TestApplyIsIdempotent();
            TestDeleteIndicesStayCorrect();

            Console.WriteLine(new string('=', 66));
            Console.ForegroundColor = _failed == 0 ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine(string.Format("עברו {0}, נכשלו {1}", _passed, _failed));
            Console.ResetColor();

            if (Failures.Count > 0)
            {
                Console.WriteLine();
                Console.WriteLine("כשלים:");
                foreach (string f in Failures) Console.WriteLine("  - " + f);
            }

            return _failed == 0 ? 0 : 1;
        }

        // ================= מסמכים מנוונים =================

        private static void TestEmptyDocument()
        {
            WithDoc(new[] { DocxBuilder.Para.Of("") }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var snapshot = adapter.ReadAll();
                var orch = new Orchestrator(adapter, new MockEngine(snapshot), rules);

                var proposal = orch.Analyze(true);
                var result = orch.Apply(proposal, null, true);

                Check("מסמך ריק אינו מפיל את המערכת", result.Success, result.ErrorMessage);
                Check("נשארה לפחות פסקה אחת", adapter.ParagraphCount >= 1);
            });
        }

        private static void TestSingleParagraph()
        {
            WithDoc(new[] { DocxBuilder.Para.Of("סימן א", true, 16, "center") }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var snapshot = adapter.ReadAll();
                var orch = new Orchestrator(adapter, new MockEngine(snapshot), rules);
                var result = orch.Apply(orch.Analyze(true), null, true);

                Check("מסמך של פסקה אחת עובד", result.Success, result.ErrorMessage);
                Check("המלל נשמר", adapter.ReadAll()[0].Text.Trim() == "סימן א");
            });
        }

        private static void TestAllEmptyParagraphs()
        {
            var paras = new List<DocxBuilder.Para>();
            for (int i = 0; i < 6; i++) paras.Add(DocxBuilder.Para.Of(""));

            WithDoc(paras, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var snapshot = adapter.ReadAll();
                var orch = new Orchestrator(adapter, new MockEngine(snapshot), rules);
                var result = orch.Apply(orch.Analyze(true), null, true);

                Check("מסמך של פסקאות ריקות בלבד עובד", result.Success, result.ErrorMessage);
                Check("לא נמחקו כל הפסקאות", adapter.ParagraphCount >= 1);
            });
        }

        private static void TestNoEmptyParagraphs()
        {
            WithDoc(new[]
            {
                DocxBuilder.Para.Of("סימן א", true, 16, "center"),
                DocxBuilder.Para.Of("גוף הטקסט כאן.", false, 12, "both")
            }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var snapshot = adapter.ReadAll();
                var orch = new Orchestrator(adapter, new MockEngine(snapshot), rules);
                var proposal = orch.Analyze(true);

                Check("אין מה למחוק כשאין פסקאות ריקות", proposal.CleanPlan.Deletes.Count == 0);
                Check("ההחלה מצליחה", orch.Apply(proposal, null, true).Success);
            });
        }

        // ================= שלד =================

        private static void TestSkeletonIsValidJson()
        {
            var paras = Paras("סימן א", "", "טקסט", "עוד טקסט");
            string sk = SkeletonBuilder.Build(paras, 0, paras.Count, new RulesFile(), null);

            bool ok = true;
            try { Core.Json.Json.ParseObject(sk); } catch { ok = false; }
            Check("השלד הוא JSON תקין", ok);
        }

        private static void TestLongParagraphTextNotSent()
        {
            string secret = "סוד" + new string('א', 400);
            var paras = Paras(secret);
            string sk = SkeletonBuilder.Build(paras, 0, paras.Count, new RulesFile(), null);

            Check("גוף פסקה ארוכה אינו נשלח", !sk.Contains("סוד"));
            Check("אורך הפסקה כן נשלח", sk.Contains("\"len\":" + secret.Length));
        }

        private static void TestPrivacyModeLeaksNothing()
        {
            var rules = new RulesFile { PrivacyMode = true };
            var paras = Paras("סימן ט", "מילה סודית מאוד");
            string sk = SkeletonBuilder.Build(paras, 0, paras.Count, rules, null);

            Check("מצב מסונן אינו שולח מילים", !sk.Contains("סודית") && !sk.Contains("מילה"));
            Check("מצב מסונן שולח תבנית", sk.Contains("pattern"));
        }

        private static void TestSkeletonHandlesEmptyWindow()
        {
            var paras = Paras("", "", "");
            bool ok = true;
            try
            {
                string sk = SkeletonBuilder.Build(paras, 0, paras.Count, new RulesFile(), null);
                Core.Json.Json.ParseObject(sk);
            }
            catch { ok = false; }
            Check("חלון שכולו פסקאות ריקות אינו מפיל את הבנייה", ok);
        }

        // ================= מנוע עוין =================

        private static DecisionSet ParseWith(string json, params string[] texts)
        {
            var paras = Paras(texts);
            var allowed = new Dictionary<int, ParagraphInfo>();
            foreach (var p in paras) allowed[p.Index] = p;
            return DecisionParser.Parse(json, allowed, new RulesFile());
        }

        private static void TestMalformedJson()
        {
            var set = ParseWith("{\"decisions\":[{\"i\":0,", "טקסט");
            Check("JSON חתוך אינו מפיל", set.Decisions.Count == 0 && set.Warnings.Count > 0);
        }

        private static void TestNotJsonAtAll()
        {
            var set = ParseWith("סליחה, לא הבנתי את הבקשה", "טקסט");
            Check("תשובה טקסטואלית נדחית", set.Decisions.Count == 0 && set.Warnings.Count > 0);
        }

        private static void TestEmptyResponse()
        {
            var set = ParseWith("", "טקסט");
            Check("תשובה ריקה נדחית", set.Decisions.Count == 0 && set.Warnings.Count > 0);
        }

        private static void TestOutOfRangeIndex()
        {
            var set = ParseWith("{\"decisions\":[{\"i\":-5,\"style\":\"כותרת 1\"},{\"i\":9999,\"style\":\"כותרת 1\"}]}", "טקסט");
            Check("אינדקסים מחוץ לטווח נדחים", set.Decisions.Count == 0);
        }

        private static void TestDuplicateDecisions()
        {
            var set = ParseWith(
                "{\"decisions\":[{\"i\":0,\"style\":\"כותרת 1\"},{\"i\":0,\"style\":\"כותרת 2\"}]}", "טקסט");

            Check("כפילות מוכרעת פעם אחת", set.Decisions.Count == 1);
            Check("ההחלטה הראשונה קובעת", set.Decisions[0].StyleName == "כותרת 1");
        }

        private static void TestDeleteEverything()
        {
            var set = ParseWith("{\"deletes\":[0,1,2]}", "אלף", "", "בית");

            Check("מחיקה מותרת רק לפסקה הריקה", set.Deletes.Count == 1 && set.Deletes[0] == 1);
            Check("שתי המחיקות האסורות נרשמו כאזהרה", set.Warnings.Count == 2);
        }

        private static void TestUnknownStyle()
        {
            var set = ParseWith("{\"decisions\":[{\"i\":0,\"style\":\"Heading 1\"}]}", "טקסט");
            Check("סגנון באנגלית שאינו ברשימה נדחה", set.Decisions.Count == 0);
        }

        private static void TestJsonWrappedInMarkdown()
        {
            string wrapped = "הנה התשובה:\n```json\n{\"decisions\":[{\"i\":0,\"style\":\"כותרת 1\"}]}\n```\nבהצלחה";
            string extracted = Core.Json.Json.ExtractObject(wrapped);

            Check("JSON עטוף ב-markdown מחולץ", extracted != null && extracted.StartsWith("{"));

            var set = ParseWith(extracted ?? "", "טקסט");
            Check("ההחלטה שחולצה תקפה", set.Decisions.Count == 1);
        }

        // ================= חוזה הבטיחות =================

        /// <summary>מתאם שמשנה את המלל בכוונה, כדי לוודא שהאימות תופס.</summary>
        private sealed class SabotageAdapter : IDocumentAdapter
        {
            private readonly IDocumentAdapter _inner;
            public bool RolledBack;

            public SabotageAdapter(IDocumentAdapter inner) { _inner = inner; }

            public int ParagraphCount { get { return _inner.ParagraphCount; } }
            public List<ParagraphInfo> ReadAll() { return _inner.ReadAll(); }
            public List<string> GetStyleNames() { return _inner.GetStyleNames(); }
            public bool EnsureStyleExists(string s) { return _inner.EnsureStyleExists(s); }
            public void SetParagraphText(int i, string t) { _inner.SetParagraphText(i, t); }
            public void DeleteParagraphs(IEnumerable<int> ix) { _inner.DeleteParagraphs(ix); }
            public void BeginBatch(string n) { _inner.BeginBatch(n); }
            public void EndBatch() { _inner.EndBatch(); }
            public void RollbackBatch() { RolledBack = true; _inner.RollbackBatch(); }
            public void Save() { _inner.Save(); }
            public void Dispose() { _inner.Dispose(); }

            /// <summary>במקום להחיל סגנון — משנה את הטקסט. בדיוק מה שאסור.</summary>
            public void ApplyStyle(int index, string styleName)
            {
                var all = _inner.ReadAll();
                foreach (var p in all)
                {
                    if (p.Index != index) continue;
                    _inner.SetParagraphText(index, p.Text + " נוסף");
                    return;
                }
            }
        }

        private static void TestSabotageIsCaughtAndRolledBack()
        {
            WithDoc(new[]
            {
                DocxBuilder.Para.Of("סימן א", true, 16, "center"),
                DocxBuilder.Para.Of("גוף הטקסט המקורי כאן ואסור שישתנה.", false, 12, "both")
            }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var snapshot = adapter.ReadAll();
                string hashBefore = TextHasher.HashDocument(snapshot);

                var sabotage = new SabotageAdapter(adapter);
                var orch = new Orchestrator(sabotage, new MockEngine(snapshot), rules);

                var proposal = orch.Analyze(snapshot, true);
                var result = orch.Apply(proposal, null, false);

                Check("שינוי מלל נתפס", !result.Success);
                Check("בוצע ביטול", sabotage.RolledBack);
                Check("אימות המלל דיווח כישלון", result.Verification != null && !result.Verification.Passed);
                Check("המסמך חזר למצבו המקורי",
                    TextHasher.HashDocument(adapter.ReadAll()) == hashBefore);
                Check("ההודעה מציינת את הפסקה",
                    result.ErrorMessage != null && result.ErrorMessage.Contains("פסקה"));
            });
        }

        private static void TestNikudChangeIsCaught()
        {
            var before = Paras("בְּרֵאשִׁית בָּרָא");
            var after = Paras("בראשית ברא");

            var v = TextVerifier.Verify(before, after);
            Check("הסרת ניקוד נתפסת כשינוי מלל", !v.Passed);
        }

        private static void TestGershayimSurvive()
        {
            var rules = new RulesFile();
            string[] cases =
            {
                "שו\"ע או\"ח סימן קפ\"א",
                "הרמב\"ם ז\"ל",
                "ס\"ק כ\"ב",
                "ר' יוחנן",
                "תשע\"ה"
            };

            bool ok = true;
            foreach (string c in cases)
                if (DeterministicCleaner.CleanText(c, rules) != c) { ok = false; break; }

            Check("ראשי תיבות וגרשיים אינם נפגעים", ok);
        }

        private static void TestCleanerIsIdempotent()
        {
            var rules = new RulesFile();
            string once = DeterministicCleaner.CleanText("  א  ב   ג .  ", rules);
            string twice = DeterministicCleaner.CleanText(once, rules);
            Check("ניקוי חוזר אינו משנה שוב", once == twice);
        }

        // ================= עיצוב פנימי בפסקה =================

        /// <summary>
        /// זה האלגוריתם המסוכן ביותר בקוד: אחרי הניקוי הטקסט נכתב חזרה
        /// לפסקה, וצריך להתחלק בין ההרצות הקיימות בלי לאבד את ההדגשות
        /// שבתוכה. כתיבה נאיבית של הפסקה כמחרוזת אחת הורסת אותן.
        /// </summary>
        private static void TestInlineBoldSurvivesCleaning()
        {
            WithDoc(new[]
            {
                DocxBuilder.Para.OfRuns(12, "both",
                    DocxBuilder.Run.Of("ובזה  יתיישב "),
                    DocxBuilder.Run.Of("כל הקושיא", true),
                    DocxBuilder.Run.Of("  שהקשו הראשונים ."))
            }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var before = adapter.ReadAll();
                string hashBefore = TextHasher.HashDocument(before);

                var orch = new Orchestrator(adapter, null, rules);
                var proposal = orch.Analyze(false);

                Check("זוהה ניקוי לביצוע", proposal.CleanPlan.Rewrites.Count == 1);

                var result = orch.Apply(proposal, new HashSet<int>(), true);
                Check("הניקוי הוחל", result.Success, result.ErrorMessage);

                var after = adapter.ReadAll();
                Check("המלל אומת", TextHasher.HashDocument(after) == hashBefore);
                Check("הרווחים הכפולים נוקו", !after[0].Text.Contains("  "));
                Check("הרווח לפני הנקודה הוסר", after[0].Text.EndsWith("הראשונים."));

                // בלי שמירה, בדיקת ה-XML הייתה קוראת את הקובץ המקורי ועוברת לשווא
                adapter.Save();

                // המבחן האמיתי: ההדגשה הפנימית עדיין קיימת, ועל אותן מילים
                int boldRuns, totalRuns;
                CountBoldRuns(doc, out boldRuns, out totalRuns);
                Check("שלוש ההרצות נשמרו", totalRuns == 3, "נמצאו " + totalRuns);
                Check("ההדגשה הפנימית שרדה את הניקוי", boldRuns == 1, "מודגשות: " + boldRuns);
                Check("הטקסט המודגש לא השתנה", BoldText(doc) == "כל הקושיא", BoldText(doc));
            });
        }

        /// <summary>רווח כפול שנמצא משני צדי גבול בין הרצות.</summary>
        private static void TestSpaceRemovedAcrossRunBoundary()
        {
            WithDoc(new[]
            {
                DocxBuilder.Para.OfRuns(12, "both",
                    DocxBuilder.Run.Of("ראשון "),
                    DocxBuilder.Run.Of(" שני", true))
            }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var before = adapter.ReadAll();
                string hashBefore = TextHasher.HashDocument(before);

                var orch = new Orchestrator(adapter, null, rules);
                var result = orch.Apply(orch.Analyze(false), new HashSet<int>(), true);

                Check("ניקוי על גבול הרצות מצליח", result.Success, result.ErrorMessage);

                var after = adapter.ReadAll();
                Check("נשאר רווח יחיד בין המילים", after[0].Text == "ראשון שני", "[" + after[0].Text + "]");
                Check("המלל אומת על גבול הרצות", TextHasher.HashDocument(after) == hashBefore);

                adapter.Save();
                Check("ההדגשה נשמרה על גבול הרצות", BoldText(doc) == "שני", BoldText(doc));
            });
        }

        private static void TestTabsAndNbspCleaned()
        {
            WithDoc(new[]
            {
                DocxBuilder.Para.Of("אלף	בית  גימל", false, 12, "both")
            }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var before = adapter.ReadAll();
                string hashBefore = TextHasher.HashDocument(before);

                var orch = new Orchestrator(adapter, null, rules);
                var result = orch.Apply(orch.Analyze(false), new HashSet<int>(), true);

                Check("ניקוי טאבים ורווחים קשיחים מצליח", result.Success, result.ErrorMessage);

                var after = adapter.ReadAll();
                Check("טאב הוחלף ברווח", !after[0].Text.Contains("	"));
                Check("המלל אומת אחרי טאבים ורווחים קשיחים",
                    TextHasher.HashDocument(after) == hashBefore);
            });
        }

        /// <summary>
        /// בודק את ההרצות ישירות מה-XML, בלי להסתמך על הקורא של המוצר.
        ///
        /// הבדיקה עוברת דרך פרסור XML ולא דרך חיפוש מחרוזות: אלמנט ריק
        /// נכתב פעם כ-&lt;w:b/&gt; ופעם כ-&lt;w:b /&gt;, וחיפוש טקסטואלי
        /// היה מדווח על היעדר הדגשה שקיימת בפועל.
        /// </summary>
        private static void CountBoldRuns(string path, out int bold, out int total)
        {
            bold = 0; total = 0;
            foreach (var run in TextRuns(path))
            {
                total++;
                if (IsBold(run)) bold++;
            }
        }

        private static string BoldText(string path)
        {
            foreach (var run in TextRuns(path))
                if (IsBold(run)) return RunText(run);
            return "(אין)";
        }

        private static readonly System.Xml.Linq.XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static List<System.Xml.Linq.XElement> TextRuns(string path)
        {
            var doc = System.Xml.Linq.XDocument.Parse(ReadDocumentXml(path));
            var list = new List<System.Xml.Linq.XElement>();
            foreach (var r in doc.Descendants(W + "r"))
            {
                var t = r.Element(W + "t");
                if (t != null && t.Value.Length > 0) list.Add(r);
            }
            return list;
        }

        private static bool IsBold(System.Xml.Linq.XElement run)
        {
            var rPr = run.Element(W + "rPr");
            return rPr != null && rPr.Element(W + "b") != null;
        }

        private static string RunText(System.Xml.Linq.XElement run)
        {
            var t = run.Element(W + "t");
            return t == null ? "" : t.Value;
        }

        private static string ReadDocumentXml(string path)
        {
            using (var zip = System.IO.Compression.ZipFile.OpenRead(path))
            {
                var entry = zip.GetEntry("word/document.xml");
                using (var sr = new StreamReader(entry.Open(), System.Text.Encoding.UTF8))
                    return sr.ReadToEnd();
            }
        }

        // ================= חלונות =================

        private static void TestWindowsCoverExactMultiple()
        {
            var rules = new RulesFile { WindowSize = 100, WindowOverlap = 0 };
            var w = WindowPlanner.Plan(300, rules);

            int max = 0;
            foreach (var win in w) if (win.End > max) max = win.End;
            Check("כפולה מדויקת מכוסה במלואה", max == 300);
        }

        private static void TestWindowsCoverOffByOne()
        {
            var rules = new RulesFile { WindowSize = 100, WindowOverlap = 10 };
            foreach (int n in new[] { 99, 100, 101, 181, 271 })
            {
                var w = WindowPlanner.Plan(n, rules);
                int max = 0;
                foreach (var win in w) if (win.End > max) max = win.End;
                if (max != n) { Check("כיסוי מלא עבור " + n + " פסקאות", false, "כוסו " + max); return; }
            }
            Check("כיסוי מלא בגדלים לא עגולים", true);
        }

        private static void TestAgreementUpgradesCertainty()
        {
            var a = new DecisionSet();
            a.Decisions.Add(new StyleDecision { Index = 3, StyleName = "כותרת 1", IsUncertain = true });

            var b = new DecisionSet();
            b.Decisions.Add(new StyleDecision { Index = 3, StyleName = "כותרת 1", IsUncertain = false });

            var m = DecisionMerger.Merge(new[] { a, b });
            Check("הסכמה בין חלונות מסירה את הספק",
                m.Decisions.Count == 1 && !m.Decisions[0].IsUncertain);
        }

        // ================= כללים =================

        private static void TestRulesRoundTrip()
        {
            var r = new RulesFile
            {
                Name = "סדרת שו\"ת",
                StructureHint = "שורה עם \"גרשיים\" ותו \\ ותו חדש\nשורה שנייה",
                PrivacyMode = true,
                MaxConsecutiveEmpty = 2,
                PrefixChars = 45
            };
            r.NeverTouchPrefixes.Add("בס\"ד");
            r.AllowedStyles = new List<string> { "כותרת 1", "כותרת 2", "רגיל" };

            var back = RulesFile.FromJson(r.ToJson());

            Check("שם נשמר", back.Name == r.Name);
            Check("הגדרה מבנית עם גרשיים ושורות נשמרת", back.StructureHint == r.StructureHint);
            Check("מצב מסונן נשמר", back.PrivacyMode);
            Check("קידומות מוגנות נשמרות",
                back.NeverTouchPrefixes.Count == 1 && back.NeverTouchPrefixes[0] == "בס\"ד");
            Check("סגנונות נשמרים", back.AllowedStyles.Count == 3);
            Check("מספרים נשמרים", back.MaxConsecutiveEmpty == 2 && back.PrefixChars == 45);
        }

        private static void TestRulesRejectsGarbage()
        {
            bool threw = false;
            try { RulesFile.FromJson("זה בכלל לא JSON"); }
            catch { threw = true; }
            Check("קובץ כללים פגום זורק חריגה ברורה", threw);
        }

        private static void TestRulesClampsInsaneValues()
        {
            var r = RulesFile.FromJson(
                "{\"prefixChars\":99999,\"windowSize\":1,\"windowOverlap\":99999}");

            Check("תחילית ארוכה מדי נחתכת", r.PrefixChars <= 200);
            Check("חלון קטן מדי מוגדל", r.WindowSize >= 50);
            Check("חפיפה אינה עולה על חצי חלון", r.WindowOverlap <= r.WindowSize / 2);
        }

        // ================= החלה =================

        private static void TestApplyOnlyCleaning()
        {
            WithDoc(new[]
            {
                DocxBuilder.Para.Of("סימן א", true, 16, "center"),
                DocxBuilder.Para.Of(""),
                DocxBuilder.Para.Of(""),
                DocxBuilder.Para.Of("טקסט  עם  רווחים  כפולים.", false, 12, "both")
            }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var snapshot = adapter.ReadAll();
                var orch = new Orchestrator(adapter, new MockEngine(snapshot), rules);
                var proposal = orch.Analyze(true);

                var result = orch.Apply(proposal, new HashSet<int>(), true);

                Check("ניקוי בלבד מצליח", result.Success, result.ErrorMessage);
                Check("לא הוחל אף סגנון", result.Report.StylesApplied == 0);
                Check("רווחים כן נוקו", result.Report.WhitespaceFixes > 0);

                foreach (var p in adapter.ReadAll())
                    if (p.StyleName != null && p.StyleName.StartsWith("כותרת"))
                    {
                        Check("לא נוספו כותרות במצב ניקוי", false);
                        return;
                    }
                Check("לא נוספו כותרות במצב ניקוי", true);
            });
        }

        private static void TestApplyIsIdempotent()
        {
            WithDoc(new[]
            {
                DocxBuilder.Para.Of("סימן א", true, 16, "center"),
                DocxBuilder.Para.Of("בענין הפסק", true, 14, "center"),
                DocxBuilder.Para.Of("גוף הטקסט.", false, 12, "both")
            }, (doc, adapter) =>
            {
                var rules = new RulesFile();

                var first = new Orchestrator(adapter, new MockEngine(adapter.ReadAll()), rules);
                first.Apply(first.Analyze(true), null, true);
                string hash1 = TextHasher.HashDocument(adapter.ReadAll());
                var styles1 = StyleList(adapter);

                var second = new Orchestrator(adapter, new MockEngine(adapter.ReadAll()), rules);
                var r2 = second.Apply(second.Analyze(true), null, true);

                Check("הרצה שנייה מצליחה", r2.Success, r2.ErrorMessage);
                Check("המלל זהה אחרי הרצה שנייה",
                    TextHasher.HashDocument(adapter.ReadAll()) == hash1);
                Check("הסגנונות זהים אחרי הרצה שנייה", styles1 == StyleList(adapter));
            });
        }

        private static void TestDeleteIndicesStayCorrect()
        {
            WithDoc(new[]
            {
                DocxBuilder.Para.Of("אלף", false, 12, "both"),
                DocxBuilder.Para.Of(""),
                DocxBuilder.Para.Of("בית", false, 12, "both"),
                DocxBuilder.Para.Of(""),
                DocxBuilder.Para.Of(""),
                DocxBuilder.Para.Of("גימל", false, 12, "both")
            }, (doc, adapter) =>
            {
                var rules = new RulesFile();
                var snapshot = adapter.ReadAll();
                var orch = new Orchestrator(adapter, new MockEngine(snapshot), rules);
                var result = orch.Apply(orch.Analyze(true), null, true);

                Check("מחיקה מרובה מצליחה", result.Success, result.ErrorMessage);

                var after = adapter.ReadAll();
                var texts = new List<string>();
                foreach (var p in after) if (!p.IsEmpty) texts.Add(p.Text.Trim());

                Check("סדר הפסקאות נשמר אחרי המחיקות",
                    texts.Count == 3 && texts[0] == "אלף" && texts[1] == "בית" && texts[2] == "גימל");
            });
        }

        // ================= עזר =================

        private static string StyleList(IDocumentAdapter adapter)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var p in adapter.ReadAll()) sb.Append(p.StyleName).Append('|');
            return sb.ToString();
        }

        private static List<ParagraphInfo> Paras(params string[] texts)
        {
            var list = new List<ParagraphInfo>();
            for (int i = 0; i < texts.Length; i++)
                list.Add(new ParagraphInfo
                {
                    Index = i,
                    Text = texts[i],
                    StyleName = "רגיל",
                    FontSize = 12,
                    Alignment = Align.Justify
                });
            return list;
        }

        private static void WithDoc(IEnumerable<DocxBuilder.Para> paras,
                                    Action<string, OpenXmlDocumentAdapter> body)
        {
            string dir = Path.Combine(Path.GetTempPath(), "sadar-qa-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "t.docx");

            try
            {
                DocxBuilder.Create(path, paras);
                using (var adapter = new OpenXmlDocumentAdapter(path))
                    body(path, adapter);
            }
            catch (Exception ex)
            {
                _failed++;
                Failures.Add("חריגה: " + ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("  נכשל  חריגה: " + ex.Message);
                Console.ResetColor();
            }
            finally
            {
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        private static void Section(string name)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("  " + name);
            Console.ResetColor();
        }

        private static void Check(string name, bool ok, string detail = "")
        {
            if (ok)
            {
                _passed++;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("    עבר   ");
            }
            else
            {
                _failed++;
                Failures.Add(name + (string.IsNullOrEmpty(detail) ? "" : " — " + detail));
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("    נכשל  ");
            }
            Console.ResetColor();
            Console.WriteLine(name + (string.IsNullOrEmpty(detail) || ok ? "" : "  — " + detail));
        }
    }
}
