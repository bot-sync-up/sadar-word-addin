using System;
using System.IO;
using System.Text;
using Sadar.Core.Rules;
using Word = Microsoft.Office.Interop.Word;

namespace Sadar.Addin
{
    /// <summary>הגדרות מקומיות. הכול נשמר אצל המשתמש בלבד; אין שרת ואין ענן.</summary>
    public static class SettingsStore
    {
        public static string Folder
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Sadar");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        public static string RulesPath { get { return Path.Combine(Folder, "כללים.json"); } }
        public static string LogPath { get { return Path.Combine(Folder, "יומן.txt"); } }

        public static RulesFile LoadRules()
        {
            try
            {
                if (File.Exists(RulesPath)) return RulesFile.Load(RulesPath);
            }
            catch { /* קובץ פגום — חוזרים לברירת מחדל במקום לחסום את המשתמש */ }

            var rules = new RulesFile();
            rules.NeverTouchPrefixes.Add("בס\"ד");
            return rules;
        }

        public static void SaveRules(RulesFile rules)
        {
            rules.Save(RulesPath);
        }

        // ---------- מנוע ----------

        public static Core.Engine.EngineSettings LoadEngine()
        {
            return Core.Engine.EngineSettings.LoadDefault();
        }

        public static void SaveEngine(Core.Engine.EngineSettings settings)
        {
            settings.SaveDefault();
        }

        public static void Log(string message)
        {
            try
            {
                File.AppendAllText(LogPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + message + Environment.NewLine,
                    new UTF8Encoding(true));
            }
            catch { }
        }
    }

    /// <summary>
    /// גיבוי והשוואה — רשת הביטחון שהמשתמש רואה בעיניים.
    /// הצעת "עבוד עם סקירת שינויים" עלתה בשרשור המקורי; כאן זה כפתור אחד.
    /// </summary>
    public static class BackupManager
    {
        public static string CreateBackup(Word.Document doc)
        {
            if (doc == null) return null;

            string full;
            try { full = doc.FullName; } catch { return null; }

            // מסמך שמעולם לא נשמר אין לו נתיב אמיתי לגבות אליו
            if (string.IsNullOrEmpty(full) || !File.Exists(full)) return null;

            string dir = Path.GetDirectoryName(full);
            string name = Path.GetFileNameWithoutExtension(full);
            string ext = Path.GetExtension(full);
            string stamp = DateTime.Now.ToString("yyyy-MM-dd_HHmm");
            string backup = Path.Combine(dir, name + ".גיבוי-לפני-סידור-" + stamp + ext);

            try
            {
                doc.Save();
                File.Copy(full, backup, true);
                SettingsStore.Log("גיבוי נוצר: " + backup);
                return backup;
            }
            catch (Exception ex)
            {
                SettingsStore.Log("יצירת גיבוי נכשלה: " + ex.Message);
                return null;
            }
        }

        public static string FindLatestBackup(Word.Document doc)
        {
            try
            {
                string full = doc.FullName;
                string dir = Path.GetDirectoryName(full);
                string name = Path.GetFileNameWithoutExtension(full);

                var files = Directory.GetFiles(dir, name + ".גיבוי-לפני-סידור-*" + Path.GetExtension(full));
                if (files.Length == 0) return null;

                Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                return files[files.Length - 1];
            }
            catch { return null; }
        }

        public static void LaunchCompare(Word.Application app, Word.Document doc)
        {
            string backup = FindLatestBackup(doc);
            if (backup == null)
                throw new InvalidOperationException(
                    "לא נמצא גיבוי למסמך הזה. גיבוי נוצר אוטומטית בכל פעם שמריצים סידור.");

            var original = app.Documents.Open(backup, ReadOnly: true, Visible: false);

            var compared = app.CompareDocuments(
                original, doc,
                Word.WdCompareDestination.wdCompareDestinationNew,
                Word.WdGranularity.wdGranularityCharLevel,
                CompareFormatting: true,
                CompareCaseChanges: true,
                CompareWhitespace: false,
                CompareComments: false);

            compared.Windows[1].Activate();

            // המרה מפורשת ל-_Document: בלעדיה יש עמימות בין המתודה Close
            // לאירוע בשם זהה, והקומפיילר בוחר בשקט
            try { ((Word._Document)original).Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
        }
    }
}
