using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Sadar.Addin;
using Sadar.Core;
using Sadar.Core.Engine;
using Sadar.Core.Model;
using Sadar.Core.Rules;
using Sadar.Core.Safety;
using Word = Microsoft.Office.Interop.Word;

namespace Sadar.WordTest
{
    /// <summary>
    /// בדיקת אינטגרציה מול וורד אמיתי.
    ///
    /// בדיקות הליבה רצות על docx ישירות ולכן אינן נוגעות בוורד כלל.
    /// כאן נבדק דווקא מה שאי אפשר לזייף: קריאת הדגשה דו-כיוונית בעברית,
    /// החלת סגנון דרך ה-Object Model, מחיקת פסקאות, רשומת ביטול יחידה,
    /// והאם המלל שרד הכל.
    ///
    /// וורד נפתח מוסתר ונסגר בסוף. אין נגיעה במסמכים פתוחים של המשתמש.
    /// </summary>
    public static class Program
    {
        private static int _passed, _failed;

        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            string path = args.Length > 0 ? Path.GetFullPath(args[0]) : null;
            if (path == null || !File.Exists(path))
            {
                Console.WriteLine("שימוש: wordtest <file.docx>");
                return 1;
            }

            Console.WriteLine("בדיקת אינטגרציה מול וורד");
            Console.WriteLine(new string('=', 60));
            Console.WriteLine("קובץ: " + Path.GetFileName(path));

            Word.Application app = null;
            Word.Document doc = null;

            try
            {
                Console.WriteLine("פותח וורד (מוסתר)...");
                app = new Word.Application();

                // הגנה: אם התחברנו בטעות למופע קיים של המשתמש, המסמכים שלו
                // נמצאים כאן. יוצאים מיד בלי לגעת בכלום ובלי לסגור שום דבר.
                int preexisting = app.Documents.Count;
                if (preexisting > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("הבדיקה בוטלה: נמצאו " + preexisting +
                        " מסמכים פתוחים במופע הזה של וורד.");
                    Console.WriteLine("סגרו את וורד והריצו שוב, כדי שהבדיקה לא תיגע במסמכים שלכם.");
                    Console.ResetColor();
                    Marshal.ReleaseComObject(app);
                    app = null;
                    return 2;
                }

                app.Visible = false;
                app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                Check("וורד נפתח במופע נקי", true);
                Console.WriteLine("גרסת וורד: " + app.Version);

                doc = app.Documents.Open(path, ReadOnly: false, Visible: false);
                Check("המסמך נפתח", doc != null);

                var adapter = new WordDocumentAdapter(app, doc);

                // ---- קריאה ----
                var before = adapter.ReadAll();
                Check("נקראו פסקאות", before.Count > 0);
                Console.WriteLine("        פסקאות: " + before.Count);

                int bold = 0, withSize = 0, centered = 0;
                foreach (var p in before)
                {
                    if (p.Bold) bold++;
                    if (p.FontSize > 0) withSize++;
                    if (p.Alignment == Align.Center) centered++;
                }

                // אלה בדיוק המאפיינים שהזיהוי נשען עליהם. אם הקריאה מהם שבורה,
                // הכל למעלה מזה יעבוד ופשוט יחזיר תוצאות שגויות.
                Check("זוהו פסקאות מודגשות", bold > 0);
                Check("נקרא גודל גופן", withSize > 0);
                Check("נקרא יישור למרכז", centered > 0);
                Console.WriteLine(string.Format("        מודגשות: {0} | עם גודל: {1} | ממורכזות: {2}",
                    bold, withSize, centered));

                string hashBefore = TextHasher.HashDocument(before);

                // ---- ניתוח והחלה ----
                var rules = new RulesFile();
                rules.NeverTouchPrefixes.Add("בס\"ד");

                var orch = new Orchestrator(adapter, new MockEngine(before), rules);
                var proposal = orch.Analyze(true);

                var changes = proposal.EffectiveChanges(rules);
                Check("הניתוח הציע שינויי סגנון", changes.Count > 0);
                Console.WriteLine("        שינויים מוצעים: " + changes.Count);
                Console.WriteLine("        ניקוי: " + proposal.CleanPlan.Rewrites.Count +
                                  " רווחים, " + proposal.CleanPlan.Deletes.Count + " פסקאות ריקות");

                int undoBefore = CountUndo(doc);

                var result = orch.Apply(proposal, null, true);
                Check("ההחלה הצליחה", result.Success,
                    result.Success ? "" : result.ErrorMessage);
                Check("אימות המלל עבר", result.Report.TextVerified);

                // ---- אימות עצמאי אחרי ----
                var after = adapter.ReadAll();
                Check("המלל לא השתנה", TextHasher.HashDocument(after) == hashBefore);

                int headings = 0;
                foreach (var p in after)
                    if (p.StyleName != null && p.StyleName.StartsWith("כותרת")) headings++;
                Check("סגנונות כותרת הוחלו בוורד", headings > 0);
                Console.WriteLine("        כותרות במסמך: " + headings);

                Check("פסקאות ריקות נמחקו", after.Count < before.Count);
                Console.WriteLine(string.Format("        פסקאות: {0} -> {1}", before.Count, after.Count));

                // ---- ביטול יחיד ----
                // כל הריצה חייבת להתבטל בלחיצת Ctrl+Z אחת. אם זה שבור,
                // המשתמש נשאר עם מסמך שלא ניתן להחזיר לאחור בקלות.
                bool undone = false;
                try { undone = doc.Undo(1); } catch { }
                Check("ביטול יחיד עבד", undone);

                if (undone)
                {
                    var reverted = adapter.ReadAll();
                    Check("הביטול החזיר את מספר הפסקאות", reverted.Count == before.Count);

                    int revertedHeadings = 0;
                    foreach (var p in reverted)
                        if (p.StyleName != null && p.StyleName.StartsWith("כותרת")) revertedHeadings++;
                    Check("הביטול הסיר את הסגנונות", revertedHeadings == 0);
                }

                ((Word._Document)doc).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                doc = null;
                Check("המסמך נסגר בלי לשמור", true);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("  נכשל  חריגה: " + ex.Message);
                Console.ResetColor();
            }
            finally
            {
                try { if (doc != null) ((Word._Document)doc).Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }

                // סוגרים את וורד רק אם לא נשאר בו שום מסמך אחר.
                // עדיף להשאיר תהליך יתום מאשר לסכן מסמך של מישהו.
                try
                {
                    if (app != null && app.Documents.Count == 0)
                        ((Word._Application)app).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }

                if (app != null) { try { Marshal.ReleaseComObject(app); } catch { } }
            }

            Console.WriteLine(new string('=', 60));
            Console.ForegroundColor = _failed == 0 ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine(string.Format("עברו {0}, נכשלו {1}", _passed, _failed));
            Console.ResetColor();
            return _failed == 0 ? 0 : 1;
        }

        private static int CountUndo(Word.Document doc)
        {
            try { return doc.Undo(0) ? 1 : 0; } catch { return 0; }
        }

        private static void Check(string name, bool ok, string detail = "")
        {
            if (ok)
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
            Console.WriteLine(name + (string.IsNullOrEmpty(detail) ? "" : "  — " + detail));
        }
    }
}
