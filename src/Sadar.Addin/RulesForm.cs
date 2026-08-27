using System;
using System.Drawing;
using System.Windows.Forms;
using Sadar.Core.Rules;

namespace Sadar.Addin
{
    /// <summary>
    /// עריכת קובץ הכללים.
    ///
    /// זה הפיצ'ר שהופך את התוסף מכלי חד-פעמי לכלי עבודה: מגדירים פעם אחת
    /// איך הסדרה בנויה, וכל ספר נוסף מסודר באותם כללים בדיוק.
    /// </summary>
    public sealed class RulesForm : Form
    {
        public RulesFile Rules { get; private set; }

        private TextBox _structure, _styles, _protect;
        private CheckBox _collapse, _trimEdges, _punct, _tabs, _removeEmpty, _trimTrailing, _privacy;
        private NumericUpDown _maxEmpty, _prefixChars, _windowSize;
        private Label _pathLabel;

        public RulesForm(RulesFile rules, string path)
        {
            Rules = rules;
            BuildUi(path);
            LoadFrom(rules);
        }

        private void BuildUi(string path)
        {
            Text = "סַדָּר — כללים";
            Width = 720;
            Height = 660;
            StartPosition = FormStartPosition.CenterParent;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(16) };
            int y = 8;

            AddHeader(scroll, "מבנה הספר", ref y);
            AddHelp(scroll,
                "תארו במילים שלכם איך הספר בנוי. זה מה שמאפשר לזהות כותרות " +
                "בלי להסתמך רק על גודל גופן והדגשה.", ref y);

            _structure = new TextBox
            {
                Multiline = true,
                Height = 84,
                Width = 640,
                Left = 12,
                Top = y,
                ScrollBars = ScrollBars.Vertical,
                RightToLeft = RightToLeft.Yes
            };
            scroll.Controls.Add(_structure);
            y += 96;

            AddHeader(scroll, "סגנונות מותרים", ref y);
            AddHelp(scroll, "שם סגנון אחד בכל שורה. המנוע לא יוכל לבחור שום סגנון אחר.", ref y);
            _styles = new TextBox
            {
                Multiline = true, Height = 84, Width = 640, Left = 12, Top = y,
                ScrollBars = ScrollBars.Vertical, RightToLeft = RightToLeft.Yes
            };
            scroll.Controls.Add(_styles);
            y += 96;

            AddHeader(scroll, "פסקאות מוגנות", ref y);
            AddHelp(scroll, "פסקה שמתחילה באחת הקידומות האלה לא תיגע כלל. קידומת אחת בכל שורה.", ref y);
            _protect = new TextBox
            {
                Multiline = true, Height = 60, Width = 640, Left = 12, Top = y,
                ScrollBars = ScrollBars.Vertical, RightToLeft = RightToLeft.Yes
            };
            scroll.Controls.Add(_protect);
            y += 72;

            AddHeader(scroll, "ניקוי", ref y);
            _collapse = AddCheck(scroll, "כיווץ רווחים כפולים", ref y);
            _trimEdges = AddCheck(scroll, "הסרת רווחים בתחילת וסוף פסקה", ref y);
            _punct = AddCheck(scroll, "הסרת רווח לפני סימני פיסוק", ref y);
            _tabs = AddCheck(scroll, "המרת טאבים לרווח", ref y);
            _removeEmpty = AddCheck(scroll, "מחיקת פסקאות ריקות עודפות", ref y);
            _trimTrailing = AddCheck(scroll, "מחיקת שורות ריקות בסוף המסמך", ref y);

            _maxEmpty = AddNumber(scroll, "שורות ריקות מותרות ברצף", 0, 5, ref y);

            AddHeader(scroll, "פרטיות ונפח", ref y);
            _privacy = AddCheck(scroll,
                "מצב מסונן — לא נשלח טקסט כלל, רק תבנית מבנית", ref y);
            AddHelp(scroll,
                "מומלץ ברשת מסוננת או בחומר רגיש. הזיהוי פחות מדויק, אבל שום מילה לא יוצאת מהמחשב.", ref y);

            _prefixChars = AddNumber(scroll, "תווים שנשלחים מכל פסקה", 20, 120, ref y);
            _windowSize = AddNumber(scroll, "פסקאות בכל חלון עיבוד", 100, 1000, ref y);

            _pathLabel = new Label
            {
                Left = 12, Top = y + 6, Width = 640, Height = 32,
                ForeColor = Color.Gray,
                Text = "הקובץ נשמר ב: " + path
            };
            scroll.Controls.Add(_pathLabel);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 56, Padding = new Padding(16, 10, 16, 10) };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };

            var ok = new Button { Text = "שמור", Width = 110, Height = 32, DialogResult = DialogResult.OK };
            ok.Click += (s, e) => SaveTo(Rules);

            var cancel = new Button { Text = "ביטול", Width = 110, Height = 32, DialogResult = DialogResult.Cancel };

            var reset = new Button { Text = "ברירת מחדל", Width = 130, Height = 32 };
            reset.Click += (s, e) => LoadFrom(new RulesFile());

            flow.Controls.Add(ok);
            flow.Controls.Add(cancel);
            flow.Controls.Add(reset);
            bottom.Controls.Add(flow);

            Controls.Add(scroll);
            Controls.Add(bottom);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        private static void AddHeader(Control parent, string text, ref int y)
        {
            parent.Controls.Add(new Label
            {
                Text = text,
                Left = 12,
                Top = y + 10,
                Width = 640,
                Height = 22,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 60, 110)
            });
            y += 36;
        }

        private static void AddHelp(Control parent, string text, ref int y)
        {
            var l = new Label
            {
                Text = text,
                Left = 12,
                Top = y,
                Width = 640,
                Height = 34,
                ForeColor = Color.FromArgb(90, 90, 90)
            };
            parent.Controls.Add(l);
            y += 38;
        }

        private static CheckBox AddCheck(Control parent, string text, ref int y)
        {
            var c = new CheckBox { Text = text, Left = 16, Top = y, Width = 620, Height = 24 };
            parent.Controls.Add(c);
            y += 26;
            return c;
        }

        private static NumericUpDown AddNumber(Control parent, string label, int min, int max, ref int y)
        {
            parent.Controls.Add(new Label { Text = label, Left = 16, Top = y + 4, Width = 300, Height = 22 });
            var n = new NumericUpDown
            {
                Left = 330, Top = y, Width = 80,
                Minimum = min, Maximum = max, RightToLeft = RightToLeft.No
            };
            parent.Controls.Add(n);
            y += 32;
            return n;
        }

        // ---------- טעינה ושמירה ----------

        private void LoadFrom(RulesFile r)
        {
            _structure.Text = r.StructureHint;
            _styles.Text = string.Join(Environment.NewLine, r.AllowedStyles.ToArray());
            _protect.Text = string.Join(Environment.NewLine, r.NeverTouchPrefixes.ToArray());

            _collapse.Checked = r.CollapseDoubleSpaces;
            _trimEdges.Checked = r.TrimParagraphEdges;
            _punct.Checked = r.RemoveSpaceBeforePunctuation;
            _tabs.Checked = r.ConvertTabsToSpaces;
            _removeEmpty.Checked = r.RemoveEmptyParagraphs;
            _trimTrailing.Checked = r.TrimTrailingEmpty;
            _privacy.Checked = r.PrivacyMode;

            _maxEmpty.Value = Clamp(r.MaxConsecutiveEmpty, _maxEmpty);
            _prefixChars.Value = Clamp(r.PrefixChars, _prefixChars);
            _windowSize.Value = Clamp(r.WindowSize, _windowSize);
        }

        private static decimal Clamp(int v, NumericUpDown n)
        {
            if (v < n.Minimum) return n.Minimum;
            if (v > n.Maximum) return n.Maximum;
            return v;
        }

        private void SaveTo(RulesFile r)
        {
            r.StructureHint = _structure.Text.Trim();

            r.AllowedStyles = SplitLines(_styles.Text);
            if (r.AllowedStyles.Count == 0)
                r.AllowedStyles = new System.Collections.Generic.List<string> { "כותרת 1", "כותרת 2", "רגיל" };

            if (!r.AllowedStyles.Contains(r.BodyStyle))
                r.AllowedStyles.Add(r.BodyStyle);

            r.NeverTouchPrefixes = SplitLines(_protect.Text);

            r.CollapseDoubleSpaces = _collapse.Checked;
            r.TrimParagraphEdges = _trimEdges.Checked;
            r.RemoveSpaceBeforePunctuation = _punct.Checked;
            r.ConvertTabsToSpaces = _tabs.Checked;
            r.RemoveEmptyParagraphs = _removeEmpty.Checked;
            r.TrimTrailingEmpty = _trimTrailing.Checked;
            r.PrivacyMode = _privacy.Checked;

            r.MaxConsecutiveEmpty = (int)_maxEmpty.Value;
            r.PrefixChars = (int)_prefixChars.Value;
            r.WindowSize = (int)_windowSize.Value;
        }

        private static System.Collections.Generic.List<string> SplitLines(string text)
        {
            var list = new System.Collections.Generic.List<string>();
            if (string.IsNullOrEmpty(text)) return list;

            foreach (string line in text.Split('\n'))
            {
                string t = line.Trim();
                if (t.Length > 0) list.Add(t);
            }
            return list;
        }
    }
}
