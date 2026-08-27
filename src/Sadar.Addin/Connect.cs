using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Sadar.Addin
{
    /// <summary>
    /// נקודת הכניסה של התוסף. וורד טוען את המחלקה הזו דרך COM.
    ///
    /// התוסף עצמו דק בכוונה: הוא מוסיף לשונית, ומעביר את כל העבודה
    /// ל-Sadar.Core. כך אותה לוגיקה בדיוק נבדקת בלי וורד בכלל.
    /// </summary>
    [ComVisible(true)]
    [Guid("7A3C1F42-9E5B-4D18-B6C4-2E9A5D3F8B71")]
    [ProgId(ProgIdName)]
    // חייב להיות AutoDispatch ולא None.
    //
    // וורד קורא לפונקציות הרצועה (onAction="OnAnalyze" וכו') דרך
    // IDispatch::GetIDsOfNames על אובייקט התוסף. עם ClassInterfaceType.None
    // ה-IDispatch הראשי הוא IDTExtensibility2, שאין בו את השמות האלה —
    // התוסף נטען, הלשונית מוצגת, וכל הכפתורים פשוט לא עושים כלום.
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        public const string ProgIdName = "Sadar.Connect";
        public const string FriendlyName = "סַדָּר — עימוד חכם";
        public const string DescriptionText = "החלת סגנונות וניקוי מסמכים, בלי לגעת במלל";

        private Word.Application _app;
        private MainForm _form;

        // ---------- מחזור חיים ----------

        public void OnConnection(object application, ext_ConnectMode connectMode,
                                 object addInInst, ref Array custom)
        {
            _app = application as Word.Application;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            CloseForm();
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { CloseForm(); }

        private void CloseForm()
        {
            if (_form == null) return;
            try { if (!_form.IsDisposed) _form.ForceClose(); } catch { }
            _form = null;
        }

        // ---------- רצועה ----------

        public string GetCustomUI(string ribbonID)
        {
            return RibbonXml;
        }

        private const string RibbonXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""SadarTab"" label=""סַדָּר"">
        <group id=""SadarMainGroup"" label=""סידור מסמך"">
          <button id=""SadarAnalyze"" label=""סדר את המסמך""
                  size=""large"" imageMso=""TableStyleGalleryWord""
                  onAction=""OnAnalyze""
                  screentip=""ניתוח המסמך והצעת סגנונות""
                  supertip=""סורק את המסמך, מזהה כותרות ומציע סגנונות. שום דבר לא מוחל לפני שתאשרו.""/>
          <button id=""SadarClean"" label=""ניקוי בלבד""
                  size=""large"" imageMso=""FormattingClear""
                  onAction=""OnClean""
                  screentip=""ניקוי רווחים ופסקאות ריקות""
                  supertip=""רץ מקומית לגמרי, בלי חיבור לאינטרנט ובלי שליחת דבר.""/>
        </group>
        <group id=""SadarSafetyGroup"" label=""בטיחות"">
          <button id=""SadarCompare"" label=""השווה לגיבוי""
                  size=""large"" imageMso=""CompareAndCombine""
                  onAction=""OnCompare""
                  screentip=""השוואה מול הגיבוי האחרון""
                  supertip=""פותח את פונקציית ההשוואה של וורד מול הגיבוי שנשמר לפני הסידור.""/>
        </group>
        <group id=""SadarSettingsGroup"" label=""הגדרות"">
          <button id=""SadarEngine"" label=""מנוע""
                  size=""large"" imageMso=""ServerConnection""
                  onAction=""OnEngine""
                  screentip=""בחירת המנוע וחיבורו""
                  supertip=""קלוד, ג'מיני או ChatGPT — דרך המנוי שלכם או דרך מפתח API.""/>
          <button id=""SadarRules"" label=""כללים""
                  size=""large"" imageMso=""RulesAndAlerts""
                  onAction=""OnRules""
                  screentip=""קובץ הכללים""
                  supertip=""הגדרת מבנה הספר, סגנונות מותרים וכללי ניקוי.""/>
          <button id=""SadarDoctor"" label=""בדיקת חיבור""
                  size=""large"" imageMso=""ServerProperties""
                  onAction=""OnDoctor""
                  screentip=""בדיקת החיבור ל-Claude Code""/>
          <button id=""SadarAbout"" label=""אודות""
                  size=""large"" imageMso=""Info""
                  onAction=""OnAbout""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        // ---------- פעולות הרצועה ----------

        public void OnAnalyze(object control) { Launch(false); }
        public void OnClean(object control) { Launch(true); }

        private void Launch(bool cleanOnly)
        {
            if (!EnsureDocument()) return;

            try
            {
                if (_form == null || _form.IsDisposed)
                    _form = new MainForm(_app);

                _form.Show();
                _form.BringToFront();
                _form.Start(cleanOnly);
            }
            catch (Exception ex)
            {
                Error("שגיאה בפתיחת סַדָּר", ex);
            }
        }

        public void OnCompare(object control)
        {
            if (!EnsureDocument()) return;
            try { BackupManager.LaunchCompare(_app, _app.ActiveDocument); }
            catch (Exception ex) { Error("שגיאה בהשוואה", ex); }
        }

        public void OnEngine(object control)
        {
            try
            {
                using (var dlg = new EngineForm(SettingsStore.LoadEngine()))
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        SettingsStore.SaveEngine(dlg.Settings);
                        Info("המנוע נשמר", "המנוע שנבחר: " + dlg.Settings.Info.DisplayName);
                    }
                }
            }
            catch (Exception ex) { Error("שגיאה בהגדרות המנוע", ex); }
        }

        public void OnRules(object control)
        {
            try
            {
                using (var dlg = new RulesForm(SettingsStore.LoadRules(), SettingsStore.RulesPath))
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                        SettingsStore.SaveRules(dlg.Rules);
                }
            }
            catch (Exception ex) { Error("שגיאה בהגדרות הכללים", ex); }
        }

        public void OnDoctor(object control)
        {
            try
            {
                var settings = SettingsStore.LoadEngine();
                Cursor.Current = Cursors.WaitCursor;

                System.Collections.Generic.List<Core.Engine.EngineFactory.EngineStatus> survey;
                try { survey = Core.Engine.EngineFactory.Survey(settings, true); }
                finally { Cursor.Current = Cursors.Default; }

                var sb = new System.Text.StringBuilder();
                sb.AppendLine("המנוע שנבחר: " + settings.Info.DisplayName);
                sb.AppendLine("מודל: " + settings.EffectiveModel);
                sb.AppendLine();

                int ready = 0;
                foreach (var st in survey)
                {
                    bool ok = st.Ready && st.Configured;
                    if (ok) ready++;
                    sb.Append(st.Info.Id == settings.EngineId ? "> " : "   ");
                    sb.Append(st.Info.DisplayName).Append(" — ");
                    sb.AppendLine(ok ? "מוכן" : (st.Problem ?? "לא מוגדר"));
                }

                sb.AppendLine();
                if (ready == 0)
                    sb.AppendLine("אף מנוע אינו מוכן. \"ניקוי בלבד\" עדיין עובד — הוא רץ מקומית.");
                else
                    sb.AppendLine(ready + " מנועים מוכנים לשימוש.");

                if (ready > 0) Info("בדיקת חיבור", sb.ToString());
                else Warn("בדיקת חיבור", sb.ToString());
            }
            catch (Exception ex) { Error("שגיאה בבדיקת החיבור", ex); }
        }

        public void OnAbout(object control)
        {
            Info("אודות סַדָּר",
                "סַדָּר — עימוד חכם למסמכי וורד" + Environment.NewLine +
                "גרסה " + Version + Environment.NewLine + Environment.NewLine +
                "התוסף מחיל סגנונות ומנקה מסמכים, ואינו נוגע במלל." + Environment.NewLine +
                "לפני כל שינוי נלקחת טביעת אצבע של הטקסט, ואחריו היא נבדקת שוב." + Environment.NewLine +
                "אם המלל השתנה ולו בתו אחד — הפעולה מתבטלת מאליה." + Environment.NewLine + Environment.NewLine +
                "המנוע הוא Claude Code המותקן במחשב שלכם, על החשבון שלכם." + Environment.NewLine +
                "אין שרת, אין מפתח, ואין איסוף נתונים." + Environment.NewLine + Environment.NewLine +
                "תוכנה חופשית — פותחה ונתרמה לקהילה.");
        }

        public const string Version = "1.0";

        // ---------- עזר ----------

        private bool EnsureDocument()
        {
            if (_app == null)
            {
                Warn("סַדָּר", "התוסף לא הצליח להתחבר לוורד. נסו להפעיל את וורד מחדש.");
                return false;
            }

            try
            {
                if (_app.Documents.Count == 0 || _app.ActiveDocument == null)
                {
                    Warn("אין מסמך פתוח", "פתחו את המסמך שברצונכם לסדר, ואז לחצו שוב.");
                    return false;
                }
            }
            catch
            {
                Warn("אין מסמך פתוח", "פתחו את המסמך שברצונכם לסדר, ואז לחצו שוב.");
                return false;
            }

            return true;
        }

        private static void Info(string title, string text)
        {
            MessageBox.Show(text, title, MessageBoxButtons.OK, MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }

        private static void Warn(string title, string text)
        {
            MessageBox.Show(text, title, MessageBoxButtons.OK, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }

        private static void Error(string title, Exception ex)
        {
            MessageBox.Show(ex.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }

        // ---------- רישום COM ----------

        private const string AddinKeyPath = @"Software\Microsoft\Office\Word\Addins\" + ProgIdName;

        [ComRegisterFunction]
        public static void Register(Type t)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(AddinKeyPath))
            {
                if (key == null) return;
                key.SetValue("FriendlyName", FriendlyName);
                key.SetValue("Description", DescriptionText);
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord); // טעינה אוטומטית בהפעלת וורד
                key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            try { Registry.CurrentUser.DeleteSubKeyTree(AddinKeyPath, false); }
            catch { }
        }
    }
}
