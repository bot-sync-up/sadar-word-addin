using System;
using System.Collections.Generic;
using System.Diagnostics;
using Sadar.Core.Document;
using Sadar.Core.Engine;
using Sadar.Core.Model;
using Sadar.Core.Rules;
using Sadar.Core.Safety;
using Sadar.Core.Scanner;
using Sadar.Core.Windowing;

namespace Sadar.Core
{
    /// <summary>הצעה מוכנה להצגה למשתמש. שום דבר ממנה עוד לא הוחל.</summary>
    public sealed class Proposal
    {
        public List<ParagraphInfo> Snapshot = new List<ParagraphInfo>();
        public DecisionSet Decisions = new DecisionSet();
        public DeterministicCleaner.CleanPlan CleanPlan = new DeterministicCleaner.CleanPlan();
        public readonly List<string> Warnings = new List<string>();
        public int WindowsProcessed;
        public TimeSpan AnalyzeDuration;

        /// <summary>ההחלטות שמשנות בפועל את הסגנון הקיים.</summary>
        public List<StyleDecision> EffectiveChanges(RulesFile rules)
        {
            var byIndex = new Dictionary<int, ParagraphInfo>();
            foreach (var p in Snapshot) byIndex[p.Index] = p;

            var list = new List<StyleDecision>();
            foreach (var d in Decisions.Decisions)
            {
                ParagraphInfo p;
                if (!byIndex.TryGetValue(d.Index, out p)) continue;
                if (string.Equals(p.StyleName, d.StyleName, StringComparison.Ordinal)) continue;
                list.Add(d);
            }
            return list;
        }
    }

    /// <summary>תוצאת החלה.</summary>
    public sealed class ApplyResult
    {
        public bool Success;
        public RunReport Report = new RunReport();
        public VerificationResult Verification;
        public string ErrorMessage;
    }

    public delegate void ProgressHandler(int current, int total, string message);

    /// <summary>
    /// מנהל הריצה. אחראי על סדר הפעולות ועל כך שהמסמך אף פעם לא נשאר במצב ביניים.
    ///
    /// הזרימה מפוצלת בכוונה לשניים:
    ///   Analyze — קורא, שואל את המנוע, מחזיר הצעה. לא נוגע במסמך כלל.
    ///   Apply   — מחיל את מה שהמשתמש אישר, מאמת, ומבטל הכול אם משהו לא תואם.
    /// </summary>
    public sealed class Orchestrator
    {
        private readonly IDocumentAdapter _doc;
        private readonly IDecisionEngine _engine;
        private readonly RulesFile _rules;

        public event ProgressHandler Progress;

        public Orchestrator(IDocumentAdapter doc, IDecisionEngine engine, RulesFile rules)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            if (rules == null) throw new ArgumentNullException("rules");
            _doc = doc;
            _engine = engine;
            _rules = rules;
        }

        // ---------- שלב א: ניתוח ----------

        public Proposal Analyze(bool useEngine)
        {
            Report(0, 1, "קורא את המסמך...");
            return Analyze(_doc.ReadAll(), useEngine);
        }

        /// <summary>
        /// ניתוח על תצלום שכבר נקרא.
        ///
        /// קיים כדי שהקריאה מהמסמך תוכל להיעשות בחוט אחד והעבודה האיטית
        /// בחוט אחר. אובייקטים של וורד שייכים לחוט שיצר אותם, ומגע בהם
        /// מחוט אחר עלול להיתקע או להיכשל באקראי — ולכן המתודה הזו
        /// אינה נוגעת ב-_doc בכלל.
        /// </summary>
        public Proposal Analyze(List<ParagraphInfo> snapshot, bool useEngine)
        {
            var sw = Stopwatch.StartNew();
            var proposal = new Proposal();
            proposal.Snapshot = snapshot;

            // הניקוי הדטרמיניסטי אינו תלוי במנוע ורץ תמיד
            proposal.CleanPlan = DeterministicCleaner.Plan(proposal.Snapshot, _rules);

            if (!useEngine || _engine == null)
            {
                proposal.AnalyzeDuration = sw.Elapsed;
                return proposal;
            }

            string unavailable = _engine.CheckAvailability();
            if (unavailable != null)
            {
                proposal.Warnings.Add(unavailable);
                proposal.AnalyzeDuration = sw.Elapsed;
                return proposal;
            }

            var windows = WindowPlanner.Plan(proposal.Snapshot.Count, _rules);
            var results = new List<DecisionSet>();
            var accumulated = new DecisionSet();

            for (int w = 0; w < windows.Count; w++)
            {
                var win = windows[w];
                Report(w, windows.Count,
                    string.Format("מנתח חלון {0} מתוך {1}...", w + 1, windows.Count));

                string context = SkeletonBuilder.BuildRunningContext(proposal.Snapshot, accumulated, _rules);
                string skeleton = SkeletonBuilder.Build(proposal.Snapshot, win.Start, win.Count, _rules, context);

                var allowed = new Dictionary<int, ParagraphInfo>();
                for (int i = win.Start; i < win.End && i < proposal.Snapshot.Count; i++)
                    allowed[proposal.Snapshot[i].Index] = proposal.Snapshot[i];

                DecisionSet set;
                try
                {
                    string raw = _engine.Decide(skeleton, _rules);
                    set = DecisionParser.Parse(raw, allowed, _rules);
                }
                catch (EngineException ex)
                {
                    set = new DecisionSet();
                    set.Warnings.Add(string.Format("חלון {0} לא סודר: {1}", w + 1, ex.Message));
                    if (ex.LooksLikeFilterBlock)
                    {
                        proposal.Warnings.AddRange(set.Warnings);
                        break; // חסימת סינון לא תיפתר בחלון הבא
                    }
                }
                catch (Exception ex)
                {
                    set = new DecisionSet();
                    set.Warnings.Add(string.Format("חלון {0} נכשל: {1}", w + 1, ex.Message));
                }

                results.Add(set);
                accumulated.Merge(set);
                proposal.WindowsProcessed++;
            }

            proposal.Decisions = DecisionMerger.Merge(results);
            proposal.Warnings.AddRange(proposal.Decisions.Warnings);
            proposal.AnalyzeDuration = sw.Elapsed;

            Report(windows.Count, windows.Count, "הניתוח הושלם.");
            return proposal;
        }

        // ---------- שלב ב: החלה ----------

        /// <param name="approvedIndexes">
        /// אינדקסי הפסקאות שהמשתמש אישר. null = הכול.
        /// </param>
        public ApplyResult Apply(Proposal proposal, HashSet<int> approvedIndexes, bool applyCleaning)
        {
            var sw = Stopwatch.StartNew();
            var result = new ApplyResult();
            var report = result.Report;
            report.ParagraphsScanned = proposal.Snapshot.Count;
            report.WindowsProcessed = proposal.WindowsProcessed;

            _doc.BeginBatch("סַדָּר — סידור מסמך");
            bool committed = false;

            try
            {
                // 1. יצירת סגנונות חסרים
                var needed = new HashSet<string>(StringComparer.Ordinal);
                foreach (var d in proposal.Decisions.Decisions)
                {
                    if (approvedIndexes != null && !approvedIndexes.Contains(d.Index)) continue;
                    needed.Add(d.StyleName);
                }
                foreach (string s in needed)
                {
                    if (!_doc.EnsureStyleExists(s))
                        report.Warnings.Add(string.Format("לא ניתן היה ליצור את הסגנון \"{0}\".", s));
                }

                // 2. ניקוי טקסט (רווחים בלבד)
                if (applyCleaning)
                {
                    foreach (var kv in proposal.CleanPlan.Rewrites)
                    {
                        _doc.SetParagraphText(kv.Key, kv.Value);
                        report.WhitespaceFixes++;
                    }
                }

                // 3. החלת סגנונות
                foreach (var d in proposal.Decisions.Decisions)
                {
                    if (approvedIndexes != null && !approvedIndexes.Contains(d.Index)) continue;

                    _doc.ApplyStyle(d.Index, d.StyleName);
                    report.StylesApplied++;
                    report.CountStyle(d.StyleName);
                    if (d.IsUncertain) report.UncertainCount++;
                }

                // 4. מחיקת פסקאות ריקות — אחרון, כי מחיקה מזיזה אינדקסים
                if (applyCleaning)
                {
                    var toDelete = new List<int>();
                    foreach (int i in proposal.CleanPlan.Deletes) toDelete.Add(i);
                    foreach (int i in proposal.Decisions.Deletes)
                        if (!toDelete.Contains(i)) toDelete.Add(i);

                    toDelete.Sort();
                    toDelete.Reverse(); // מהסוף להתחלה

                    if (toDelete.Count > 0)
                    {
                        _doc.DeleteParagraphs(toDelete);
                        report.ParagraphsDeleted = toDelete.Count;
                    }
                }

                // 5. אימות המלל — נקודת האל-חזור
                Report(1, 1, "מאמת שהמלל לא השתנה...");
                var after = _doc.ReadAll();
                var verification = TextVerifier.Verify(proposal.Snapshot, after);
                result.Verification = verification;

                if (!verification.Passed)
                {
                    _doc.RollbackBatch();
                    result.Success = false;
                    result.ErrorMessage = verification.Message;
                    report.TextVerified = false;
                    return result;
                }

                _doc.EndBatch();
                committed = true;

                report.TextVerified = true;
                report.Warnings.AddRange(proposal.Warnings);
                report.Duration = sw.Elapsed;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                if (!committed)
                {
                    try { _doc.RollbackBatch(); }
                    catch (Exception rollbackEx)
                    {
                        result.ErrorMessage = "שגיאה בהחלה, וגם הביטול נכשל: " + rollbackEx.Message +
                                              " | השגיאה המקורית: " + ex.Message;
                        result.Success = false;
                        return result;
                    }
                }
                result.Success = false;
                result.ErrorMessage = "הפעולה בוטלה עקב שגיאה: " + ex.Message;
                return result;
            }
        }

        private void Report(int current, int total, string message)
        {
            var h = Progress;
            if (h != null) h(current, total, message);
        }
    }
}
