using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Sadar.Core.Engine;

namespace Sadar.Addin
{
    /// <summary>
    /// בחירת המנוע וחיבורו.
    ///
    /// שני מסלולים לכל ספק: דרך המנוי שכבר יש למשתמש, או דרך מפתח API.
    /// המסלול הראשון עדיף כמעט תמיד — הוא לא עולה כסף נוסף — ולכן
    /// המנועים של המנוי מוצגים ראשונים.
    /// </summary>
    public sealed class EngineForm : Form
    {
        public EngineSettings Settings { get; private set; }

        private ListView _list;
        private TextBox _key, _model, _baseUrl;
        private Label _hint, _status;
        private Button _test, _clearKey;
        private CheckBox _showKey;
        private Thread _worker;

        public EngineForm(EngineSettings settings)
        {
            Settings = settings ?? new EngineSettings();
            BuildUi();
            LoadList();
        }

        private void BuildUi()
        {
            Text = "סַדָּר — מנוע";
            Width = 780;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            var header = new Panel { Dock = DockStyle.Top, Height = 58, Padding = new Padding(16, 12, 16, 4) };
            header.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                Text = "בחרו במה להשתמש לזיהוי הכותרות. אפשר להשתמש במנוי שכבר יש לכם, " +
                       "או במפתח API. הניקוי עובד מקומית בכל מקרה.",
                ForeColor = Color.FromArgb(80, 80, 80)
            });

            _list = new ListView
            {
                Dock = DockStyle.Top,
                Height = 210,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true
            };
            _list.Columns.Add("מנוע", 210);
            _list.Columns.Add("ספק", 100);
            _list.Columns.Add("חיבור", 110);
            _list.Columns.Add("מצב", 310);
            _list.SelectedIndexChanged += (s, e) => OnEngineSelected();

            var body = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16, 10, 16, 6) };
            int y = 6;

            _hint = new Label
            {
                Left = 8, Top = y, Width = 720, Height = 34,
                ForeColor = Color.FromArgb(30, 60, 110),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            body.Controls.Add(_hint);
            y += 40;

            body.Controls.Add(new Label { Text = "מפתח API", Left = 8, Top = y + 4, Width = 90 });
            _key = new TextBox
            {
                Left = 104, Top = y, Width = 480,
                UseSystemPasswordChar = true,
                RightToLeft = RightToLeft.No
            };
            body.Controls.Add(_key);

            _showKey = new CheckBox { Text = "הצג", Left = 596, Top = y + 2, Width = 60 };
            _showKey.CheckedChanged += (s, e) => _key.UseSystemPasswordChar = !_showKey.Checked;
            body.Controls.Add(_showKey);

            _clearKey = new Button { Text = "מחק", Left = 660, Top = y - 1, Width = 68, Height = 25 };
            _clearKey.Click += (s, e) => { _key.Text = ""; };
            body.Controls.Add(_clearKey);
            y += 34;

            body.Controls.Add(new Label
            {
                Text = "המפתח נשמר מוצפן למשתמש הזה בלבד, ואינו קריא ממחשב אחר.",
                Left = 104, Top = y, Width = 620, Height = 18,
                ForeColor = Color.Gray
            });
            y += 28;

            body.Controls.Add(new Label { Text = "מודל", Left = 8, Top = y + 4, Width = 90 });
            _model = new TextBox { Left = 104, Top = y, Width = 300, RightToLeft = RightToLeft.No };
            body.Controls.Add(_model);
            body.Controls.Add(new Label
            {
                Text = "ריק = ברירת המחדל של המנוע",
                Left = 412, Top = y + 4, Width = 300, ForeColor = Color.Gray
            });
            y += 34;

            body.Controls.Add(new Label { Text = "כתובת בסיס", Left = 8, Top = y + 4, Width = 90 });
            _baseUrl = new TextBox { Left = 104, Top = y, Width = 480, RightToLeft = RightToLeft.No };
            body.Controls.Add(_baseUrl);
            body.Controls.Add(new Label
            {
                Text = "לשער חלופי",
                Left = 592, Top = y + 4, Width = 140, ForeColor = Color.Gray
            });
            y += 40;

            _test = new Button { Text = "בדוק חיבור", Left = 8, Top = y, Width = 130, Height = 30 };
            _test.Click += (s, e) => TestConnection();
            body.Controls.Add(_test);

            _status = new Label
            {
                Left = 148, Top = y + 4, Width = 580, Height = 60,
                ForeColor = Color.FromArgb(60, 60, 60)
            };
            body.Controls.Add(_status);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 54, Padding = new Padding(16, 10, 16, 10) };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };

            var ok = new Button { Text = "שמור", Width = 110, Height = 30, DialogResult = DialogResult.OK };
            ok.Click += (s, e) => Commit();

            var cancel = new Button { Text = "ביטול", Width = 110, Height = 30, DialogResult = DialogResult.Cancel };

            flow.Controls.Add(ok);
            flow.Controls.Add(cancel);
            bottom.Controls.Add(flow);

            Controls.Add(body);
            Controls.Add(_list);
            Controls.Add(header);
            Controls.Add(bottom);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        // ---------- רשימת המנועים ----------

        private void LoadList()
        {
            _list.Items.Clear();

            foreach (var info in EngineInfo.All)
            {
                var item = new ListViewItem(info.DisplayName);
                item.SubItems.Add(info.Vendor);
                item.SubItems.Add(info.NeedsApiKey ? "מפתח API" : "מנוי");
                item.SubItems.Add(DescribeState(info));
                item.Tag = info;

                if (info.Id == Settings.EngineId)
                {
                    item.Selected = true;
                    item.Font = new Font(_list.Font, FontStyle.Bold);
                }

                _list.Items.Add(item);
            }

            OnEngineSelected();
        }

        /// <summary>
        /// מצב מהיר בלי גישה לרשת. בדיקה אמיתית נעשית רק בלחיצה על "בדוק חיבור",
        /// כדי שפתיחת החלון לא תיתקע על שבע קריאות רשת.
        /// </summary>
        private string DescribeState(EngineInfo info)
        {
            if (info.NeedsApiKey)
            {
                string hint = Settings.KeyHint(info.Id);
                return hint == null ? "לא הוגדר מפתח" : "מפתח שמור " + hint;
            }

            try
            {
                var probe = new EngineSettings { EngineId = info.Id };
                var engine = EngineFactory.Create(probe);

                var cli = engine as CliEngineBase;
                if (cli != null)
                    return string.IsNullOrEmpty(cli.Path) ? "לא מותקן" : "מותקן";

                var claude = engine as ClaudeCodeEngine;
                if (claude != null)
                    return string.IsNullOrEmpty(claude.ExePath) ? "לא מותקן" : "מותקן";
            }
            catch { }

            return "";
        }

        private EngineInfo Selected
        {
            get
            {
                if (_list.SelectedItems.Count == 0) return EngineInfo.All[0];
                return _list.SelectedItems[0].Tag as EngineInfo ?? EngineInfo.All[0];
            }
        }

        private void OnEngineSelected()
        {
            var info = Selected;
            bool needsKey = info.NeedsApiKey;

            _key.Enabled = needsKey;
            _showKey.Enabled = needsKey;
            _clearKey.Enabled = needsKey;

            _key.Text = needsKey ? (Settings.GetKey(info.Id) ?? "") : "";
            _baseUrl.Enabled = needsKey;

            _model.Text = info.Id == Settings.EngineId && !string.IsNullOrEmpty(Settings.Model)
                ? Settings.Model : "";
            _model.ForeColor = Color.Black;

            _baseUrl.Text = info.Id == Settings.EngineId ? (Settings.BaseUrl ?? "") : "";

            _hint.Text = needsKey
                ? "מפתח: " + info.HowToConnect
                : "התקנה: " + info.HowToConnect;

            _status.Text = "מודל ברירת מחדל: " + info.DefaultModel;
            _status.ForeColor = Color.FromArgb(60, 60, 60);
        }

        // ---------- בדיקה ----------

        private void TestConnection()
        {
            if (_worker != null && _worker.IsAlive) return;

            var probe = BuildProbe();

            _test.Enabled = false;
            _status.ForeColor = Color.FromArgb(60, 60, 60);
            _status.Text = "בודק...";

            _worker = new Thread(() =>
            {
                string problem;
                try
                {
                    var engine = EngineFactory.Create(probe);
                    problem = engine.CheckAvailability();
                }
                catch (Exception ex) { problem = ex.Message; }

                BeginInvoke(new Action(() =>
                {
                    _test.Enabled = true;
                    if (problem == null)
                    {
                        _status.ForeColor = Color.FromArgb(20, 120, 40);
                        _status.Text = "החיבור תקין. המנוע מוכן לשימוש.";
                    }
                    else
                    {
                        _status.ForeColor = Color.FromArgb(170, 40, 40);
                        _status.Text = problem;
                    }
                }));
            });
            _worker.IsBackground = true;
            _worker.Start();
        }

        private EngineSettings BuildProbe()
        {
            var info = Selected;

            var probe = new EngineSettings
            {
                EngineId = info.Id,
                Model = _model.Text.Trim(),
                BaseUrl = _baseUrl.Text.Trim(),
                TimeoutSeconds = Settings.TimeoutSeconds,
                CliPath = Settings.CliPath
            };

            if (info.NeedsApiKey)
            {
                string typed = _key.Text.Trim();
                probe.SetKey(info.Id, string.IsNullOrEmpty(typed) ? Settings.GetKey(info.Id) : typed);
            }

            return probe;
        }

        // ---------- שמירה ----------

        private void Commit()
        {
            var info = Selected;

            Settings.EngineId = info.Id;
            Settings.Model = _model.Text.Trim();
            Settings.BaseUrl = _baseUrl.Text.Trim();

            if (info.NeedsApiKey)
            {
                string typed = _key.Text.Trim();
                // שדה ריק כשכבר יש מפתח שמור אינו מחיקה — מחיקה נעשית בכפתור
                if (!string.IsNullOrEmpty(typed) || !Settings.HasKey(info.Id))
                    Settings.SetKey(info.Id, typed);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // מפתח שהוקלד ולא נשמר לא צריך להישאר בזיכרון התהליך
            _key.Text = new string('\0', 1);
            _key.Clear();
            base.OnFormClosing(e);
        }
    }
}
