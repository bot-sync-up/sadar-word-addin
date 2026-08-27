using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Sadar.Core;
using Sadar.Core.Engine;
using Sadar.Core.Model;
using Sadar.Core.Rules;
using Word = Microsoft.Office.Interop.Word;

namespace Sadar.Addin
{
    /// <summary>
    /// החלון הראשי: מציג את ההצעה לפני שמשהו מוחל.
    ///
    /// זו הנקודה שבה המשתמש שומר על השליטה. אפשר לראות כל החלטה,
    /// לקפוץ לפסקה במסמך, ולבטל שורות בודדות. פסקאות שהמנוע לא היה
    /// בטוח לגביהן צפות לראש הרשימה.
    /// </summary>
    public sealed class MainForm : Form
    {
        private readonly Word.Application _app;

        private DataGridView _grid;
        private Label _status;
        private ProgressBar _progress;
        private Button _btnApply, _btnCancel, _btnSelectAll, _btnSelectNone;
        private CheckBox _chkClean;
        private Label _summary;

        private Proposal _proposal;
        private RulesFile _rules;
        private WordDocumentAdapter _adapter;
        private Orchestrator _orchestrator;
        private Thread _worker;
        private bool _cleanOnly;
        private bool _populating;
        private bool _allowClose;

        private const int ColApprove = 0, ColPara = 1, ColText = 2, ColFrom = 3, ColTo = 4, ColNote = 5;

        public MainForm(Word.Application app)
        {
            _app = app;
            BuildUi();
        }

        // ---------- ממשק ----------

        private void BuildUi()
        {
            Text = "סַדָּר — עימוד חכם";
            Width = 980;
            Height = 640;
            StartPosition = FormStartPosition.CenterScreen;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            Font = new Font("Segoe UI", 9.5f);
            MinimumSize = new Size(760, 480);

            var top = new Panel { Dock = DockStyle.Top, Height = 64, Padding = new Padding(12, 10, 12, 6) };

            _status = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                Text = "מוכן.",
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleRight
            };

            _progress = new ProgressBar { Dock = DockStyle.Top, Height = 16, Style = ProgressBarStyle.Continuous };

            top.Controls.Add(_progress);
            top.Controls.Add(_status);

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                EditMode = DataGridViewEditMode.EditOnEnter,
                RightToLeft = RightToLeft.Yes,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None
            };

            _grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "החל", Width = 46 });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "פסקה", Width = 62, ReadOnly = true });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "תחילת הטקסט", Width = 380, ReadOnly = true });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "סגנון נוכחי", Width = 110, ReadOnly = true });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "סגנון מוצע", Width = 110, ReadOnly = true });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "הערה", Width = 220, ReadOnly = true });

            _grid.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) JumpToParagraph(e.RowIndex); };
            _grid.SelectionChanged += (s, e) =>
            {
                // מילוי הטבלה משנה את הבחירה בכל שורה שנוספת. בלי השמירה הזו
                // כל שורה הייתה גוררת קפיצה במסמך — מאות קריאות COM מיותרות.
                if (_populating) return;
                if (_grid.CurrentRow != null) JumpToParagraph(_grid.CurrentRow.Index);
            };

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 84, Padding = new Padding(12, 8, 12, 10) };

            _summary = new Label
            {
                Dock = DockStyle.Top,
                Height = 34,
                TextAlign = ContentAlignment.MiddleRight,
                ForeColor = Color.FromArgb(70, 70, 70)
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };

            _btnApply = MakeButton("החל את המסומנים", 160, true);
            _btnApply.Click += OnApplyClick;

            _btnCancel = MakeButton("סגור", 90, false);
            _btnCancel.Click += (s, e) => Close();

            _btnSelectAll = MakeButton("סמן הכל", 100, false);
            _btnSelectAll.Click += (s, e) => SetAllChecks(true);

            _btnSelectNone = MakeButton("נקה סימון", 100, false);
            _btnSelectNone.Click += (s, e) => SetAllChecks(false);

            _chkClean = new CheckBox
            {
                Text = "בצע גם ניקוי רווחים ופסקאות ריקות",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(18, 8, 6, 0)
            };

            buttons.Controls.Add(_btnApply);
            buttons.Controls.Add(_btnCancel);
            buttons.Controls.Add(_btnSelectAll);
            buttons.Controls.Add(_btnSelectNone);
            buttons.Controls.Add(_chkClean);

            bottom.Controls.Add(buttons);
            bottom.Controls.Add(_summary);

            Controls.Add(_grid);
            Controls.Add(bottom);
            Controls.Add(top);

            SetBusy(false);
        }

        private static Button MakeButton(string text, int width, bool primary)
        {
            var b = new Button
            {
                Text = text,
                Width = width,
                Height = 32,
                Margin = new Padding(6, 4, 6, 4),
                UseVisualStyleBackColor = true
            };
            if (primary) b.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            return b;
        }

        // ---------- ריצה ----------

        public void Start(bool cleanOnly)
        {
            if (_worker != null && _worker.IsAlive) return;

            _cleanOnly = cleanOnly;
            _grid.Rows.Clear();
            _rules = SettingsStore.LoadRules();

            Word.Document doc;
            try { doc = _app.ActiveDocument; }
            catch
            {
                ShowStatus("אין מסמך פתוח.", true);
                return;
            }

            _adapter = new WordDocumentAdapter(_app, doc);

            IDecisionEngine engine = cleanOnly ? null : new ClaudeCodeEngine();
            _orchestrator = new Orchestrator(_adapter, engine, _rules);
            _orchestrator.Progress += OnProgress;

            SetBusy(true);
            ShowStatus(cleanOnly ? "מנקה..." : "קורא את המסמך...", false);

            // הקריאה מהמסמך נעשית כאן, על החוט של וורד. אובייקטים של וורד
            // שייכים לחוט שיצר אותם, ומגע בהם מחוט אחר נכשל באקראי או נתקע.
            List<ParagraphInfo> snapshot;
            try
            {
                snapshot = _adapter.ReadAll();
            }
            catch (Exception ex)
            {
                SetBusy(false);
                ShowStatus("קריאת המסמך נכשלה: " + ex.Message, true);
                return;
            }

            // מכאן והלאה רק עבודה על נתונים — בלי שום נגיעה בוורד.
            _worker = new Thread(() =>
            {
                try
                {
                    var proposal = _orchestrator.Analyze(snapshot, !cleanOnly);
                    BeginInvoke(new Action(() => OnAnalyzeDone(proposal)));
                }
                catch (Exception ex)
                {
                    BeginInvoke(new Action(() =>
                    {
                        SetBusy(false);
                        ShowStatus("הניתוח נכשל: " + ex.Message, true);
                        SettingsStore.Log("ניתוח נכשל: " + ex);
                    }));
                }
            });
            _worker.SetApartmentState(ApartmentState.STA);
            _worker.IsBackground = true;
            _worker.Start();
        }

        private void OnProgress(int current, int total, string message)
        {
            if (InvokeRequired) { BeginInvoke(new Action(() => OnProgress(current, total, message))); return; }

            _status.Text = message;
            if (total > 0)
            {
                _progress.Maximum = Math.Max(1, total);
                _progress.Value = Math.Min(current, _progress.Maximum);
            }
        }

        private void OnAnalyzeDone(Proposal proposal)
        {
            _proposal = proposal;
            SetBusy(false);
            PopulateGrid();

            foreach (string w in proposal.Warnings) SettingsStore.Log("אזהרה: " + w);

            if (proposal.Warnings.Count > 0 && proposal.Decisions.Count == 0 && !_cleanOnly)
            {
                ShowStatus(proposal.Warnings[0], true);
                return;
            }

            ShowStatus("הניתוח הושלם. עברו על ההצעה ואשרו.", false);
        }

        private void PopulateGrid()
        {
            _populating = true;
            try { PopulateGridCore(); }
            finally { _populating = false; }
        }

        private void PopulateGridCore()
        {
            _grid.Rows.Clear();
            if (_proposal == null) return;

            var byIndex = new Dictionary<int, ParagraphInfo>();
            foreach (var p in _proposal.Snapshot) byIndex[p.Index] = p;

            var changes = _proposal.EffectiveChanges(_rules);

            // הלא-בטוחות קודם — הן אלה שדורשות עין אנושית
            changes.Sort((a, b) =>
            {
                if (a.IsUncertain != b.IsUncertain) return a.IsUncertain ? -1 : 1;
                return a.Index.CompareTo(b.Index);
            });

            foreach (var d in changes)
            {
                ParagraphInfo p;
                byIndex.TryGetValue(d.Index, out p);

                int row = _grid.Rows.Add(
                    true,
                    d.Index + 1,
                    p == null ? "" : p.Prefix(90),
                    p == null ? "" : p.StyleName,
                    d.StyleName,
                    d.IsUncertain ? d.Reason : "");

                if (d.IsUncertain)
                {
                    _grid.Rows[row].DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 225);
                    _grid.Rows[row].DefaultCellStyle.ForeColor = Color.FromArgb(120, 80, 0);
                }

                _grid.Rows[row].Tag = d.Index;
            }

            int uncertain = 0;
            foreach (var d in changes) if (d.IsUncertain) uncertain++;

            _summary.Text = string.Format(
                "{0} שינויי סגנון · {1} טעונות בדיקה · {2} תיקוני רווחים · {3} פסקאות ריקות למחיקה · {4} פסקאות במסמך",
                changes.Count, uncertain,
                _proposal.CleanPlan.Rewrites.Count,
                _proposal.CleanPlan.Deletes.Count,
                _proposal.Snapshot.Count);

            _btnApply.Enabled = changes.Count > 0 || _proposal.CleanPlan.TotalChanges > 0;
        }

        private void JumpToParagraph(int rowIndex)
        {
            if (_proposal == null || rowIndex < 0 || rowIndex >= _grid.Rows.Count) return;

            var tag = _grid.Rows[rowIndex].Tag;
            if (!(tag is int)) return;
            int paraIndex = (int)tag;

            try
            {
                var doc = _app.ActiveDocument;
                if (paraIndex < 0 || paraIndex >= doc.Paragraphs.Count) return;
                doc.Paragraphs[paraIndex + 1].Range.Select();
                _app.ActiveWindow.ScrollIntoView(doc.Paragraphs[paraIndex + 1].Range);
            }
            catch { /* המשתמש סגר או החליף מסמך */ }
        }

        private void SetAllChecks(bool value)
        {
            foreach (DataGridViewRow row in _grid.Rows)
                row.Cells[ColApprove].Value = value;
        }

        // ---------- החלה ----------

        private void OnApplyClick(object sender, EventArgs e)
        {
            if (_proposal == null) return;

            var approved = new HashSet<int>();
            foreach (DataGridViewRow row in _grid.Rows)
            {
                object v = row.Cells[ColApprove].Value;
                if (v is bool && (bool)v && row.Tag is int) approved.Add((int)row.Tag);
            }

            bool clean = _chkClean.Checked;

            if (approved.Count == 0 && !clean)
            {
                ShowStatus("לא נבחר דבר להחלה.", true);
                return;
            }

            string backup = null;
            try { backup = BackupManager.CreateBackup(_app.ActiveDocument); }
            catch (Exception ex) { SettingsStore.Log("גיבוי נכשל: " + ex.Message); }

            if (backup == null)
            {
                var answer = MessageBox.Show(
                    "לא ניתן היה ליצור גיבוי — כנראה שהמסמך עדיין לא נשמר לדיסק." + Environment.NewLine +
                    Environment.NewLine +
                    "אפשר להמשיך: כל השינויים ניתנים לביטול ב-Ctrl+Z, והמלל מאומת אוטומטית." +
                    Environment.NewLine + "להמשיך בכל זאת?",
                    "אין גיבוי", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2,
                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);

                if (answer != DialogResult.Yes) return;
            }

            SetBusy(true);
            ShowStatus("מחיל ומאמת...", false);

            ApplyResult result;
            try
            {
                result = _orchestrator.Apply(_proposal, approved, clean);
            }
            catch (Exception ex)
            {
                SetBusy(false);
                ShowStatus("ההחלה נכשלה: " + ex.Message, true);
                SettingsStore.Log("החלה נכשלה: " + ex);
                return;
            }

            SetBusy(false);
            result.Report.BackupPath = backup;

            if (result.Success) ShowSuccess(result);
            else ShowFailure(result);
        }

        private void ShowSuccess(ApplyResult result)
        {
            var r = result.Report;
            ShowStatus("הושלם. המלל אומת ולא השתנה.", false);
            SettingsStore.Log(string.Format(
                "הצלחה: {0} סגנונות, {1} רווחים, {2} פסקאות נמחקו", r.StylesApplied, r.WhitespaceFixes, r.ParagraphsDeleted));

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("הסידור הושלם.");
            sb.AppendLine();
            sb.AppendLine("סגנונות שהוחלו: " + r.StylesApplied);
            foreach (var kv in r.StyleCounts) sb.AppendLine("    " + kv.Key + ": " + kv.Value);
            sb.AppendLine("תיקוני רווחים: " + r.WhitespaceFixes);
            sb.AppendLine("פסקאות ריקות שנמחקו: " + r.ParagraphsDeleted);
            if (r.UncertainCount > 0) sb.AppendLine("הוחלו למרות שסומנו לבדיקה: " + r.UncertainCount);
            sb.AppendLine();
            sb.AppendLine("המלל אומת — לא השתנה בו דבר.");
            sb.AppendLine();
            if (r.BackupPath != null)
                sb.AppendLine("גיבוי: " + System.IO.Path.GetFileName(r.BackupPath));
            sb.AppendLine("לביטול מלא: Ctrl+Z");

            MessageBox.Show(sb.ToString(), "סַדָּר", MessageBoxButtons.OK, MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);

            _grid.Rows.Clear();
            _btnApply.Enabled = false;
        }

        private void ShowFailure(ApplyResult result)
        {
            ShowStatus(result.ErrorMessage, true);
            SettingsStore.Log("בוטל: " + result.ErrorMessage);

            MessageBox.Show(
                result.ErrorMessage + Environment.NewLine + Environment.NewLine +
                "המסמך הוחזר למצבו הקודם. לא נשאר שינוי חלקי.",
                "הפעולה בוטלה", MessageBoxButtons.OK, MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }

        // ---------- מצב ----------

        private void SetBusy(bool busy)
        {
            _progress.Visible = busy;
            _progress.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
            _btnApply.Enabled = !busy && _grid.Rows.Count > 0;
            _btnSelectAll.Enabled = !busy;
            _btnSelectNone.Enabled = !busy;
            _grid.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        private void ShowStatus(string text, bool isProblem)
        {
            _status.Text = text;
            _status.ForeColor = isProblem ? Color.FromArgb(180, 30, 30) : Color.FromArgb(20, 20, 20);
        }

        /// <summary>סגירה אמיתית, לשימוש בעת כיבוי וורד.</summary>
        public void ForceClose()
        {
            _allowClose = true;
            Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_worker != null && _worker.IsAlive)
            {
                var answer = MessageBox.Show(
                    "הניתוח עדיין רץ. לסגור בכל זאת?", "סַדָּר",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2,
                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);

                if (answer != DialogResult.Yes) { e.Cancel = true; return; }
            }

            // בשגרה מסתירים במקום להשמיד — התוסף חי כל עוד וורד חי.
            // אבל כשוורד נסגר צריך לאפשר סגירה אמיתית, אחרת החלון
            // לעולם לא משוחרר והתהליך עלול להישאר תלוי.
            if (!_allowClose && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
            }
            base.OnFormClosing(e);
        }
    }
}
