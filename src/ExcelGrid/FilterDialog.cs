using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelGrid.Ssms;

internal sealed class FilterDialog : Form
{
    private readonly ComboBox _operator = new();
    private readonly TextBox _operand = new();
    private readonly TextBox _search = new();
    private readonly FlowLayoutPanel _valueList = new();
    private readonly Label _selectionLabel = new();
    private Button? _applyButton;
    private readonly IReadOnlyList<string> _allValues;
    private readonly HashSet<string> _selectedValues;
    private readonly Dictionary<string, CheckBox> _checks = new(StringComparer.CurrentCultureIgnoreCase);
    private readonly Palette _palette;
    private readonly bool _truncated;

    public FilterSpec? Result { get; private set; }

    public FilterDialog(string columnName, IReadOnlyList<string> values, bool truncated, FilterSpec? current, bool dark, Point popupLocation)
    {
        _allValues = values;
        _truncated = truncated;
        _selectedValues = current?.AllowedValues != null
            ? new HashSet<string>(current.AllowedValues, StringComparer.CurrentCultureIgnoreCase)
            : new HashSet<string>(values, StringComparer.CurrentCultureIgnoreCase);
        _palette = dark ? Palette.Dark : Palette.Light;

        Text = "Filter " + columnName;
        StartPosition = FormStartPosition.Manual;
        Location = ClampToWorkingArea(popupLocation, Size);
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        MinimizeBox = false;
        MaximizeBox = false;
        MinimumSize = new Size(420, 500);
        Size = new Size(440, 590);
        Padding = new Padding(1);
        BackColor = _palette.Border;
        Font = new Font("Segoe UI", 9.5f);

        Controls.Add(BuildBody(columnName, current));
        Shown += (_, _) =>
        {
            Location = ClampToWorkingArea(popupLocation, Size);
            if (_applyButton != null)
            {
                _applyButton.Enabled = true;
                _applyButton.BringToFront();
            }
            SetCueBanner(_search, "Search values");
            SetCueBanner(_operand, "Value");
            _search.Focus();
        };
        Resize += (_, _) => ApplyRoundedRegion();
    }

    private static Point ClampToWorkingArea(Point desired, Size size)
    {
        var area = Screen.FromPoint(desired).WorkingArea;
        var x = Math.Max(area.Left, Math.Min(desired.X, area.Right - size.Width));
        var y = desired.Y + size.Height <= area.Bottom ? desired.Y : Math.Max(area.Top, desired.Y - size.Height - 6);
        return new Point(x, y);
    }

    protected override CreateParams CreateParams
    {
        get
        {
            const int CsDropShadow = 0x00020000;
            var parameters = base.CreateParams;
            parameters.ClassStyle |= CsDropShadow;
            return parameters;
        }
    }

    private Control BuildBody(string columnName, FilterSpec? current)
    {
        var body = new Panel { Dock = DockStyle.Fill, BackColor = _palette.Surface };
        body.Controls.Add(BuildValues());
        body.Controls.Add(BuildFooter());
        body.Controls.Add(BuildValueToolbar());
        body.Controls.Add(BuildSearch());
        body.Controls.Add(BuildCondition(current));
        body.Controls.Add(BuildTitle(columnName));
        return body;
    }

    private Control BuildTitle(string columnName)
    {
        var panel = new Panel { Dock = DockStyle.Top, Height = 66, Padding = new Padding(18, 11, 10, 8), BackColor = _palette.Surface };
        var version = typeof(FilterDialog).Assembly.GetName().Version;
        var title = new Label { Text = $"Excel Grid v{version?.Major}.{version?.Minor}.{version?.Build}", AutoSize = true, ForeColor = _palette.Muted, Font = new Font(Font.FontFamily, 8.5f), Location = new Point(18, 10) };
        var column = new Label { Text = columnName, AutoEllipsis = true, ForeColor = _palette.Text, Font = new Font(Font.FontFamily, 13f, FontStyle.Bold), Location = new Point(18, 29), Size = new Size(345, 27) };
        var close = MakeButton("×", _palette.Surface, _palette.Text, 36, 34);
        close.Font = new Font(Font.FontFamily, 14f);
        close.Location = new Point(0, 9);
        close.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
        panel.Controls.Add(title);
        panel.Controls.Add(column);
        panel.Controls.Add(close);
        panel.Resize += (_, _) => close.Location = new Point(panel.ClientSize.Width - close.Width - 10, 9);
        panel.MouseDown += DragWindow;
        title.MouseDown += DragWindow;
        column.MouseDown += DragWindow;
        return panel;
    }

    private Control BuildCondition(FilterSpec? current)
    {
        var card = new Panel { Dock = DockStyle.Top, Height = 92, Padding = new Padding(18, 7, 18, 10), BackColor = _palette.Card };
        var label = new Label { Text = "TEXT CONDITION", Dock = DockStyle.Top, Height = 22, ForeColor = _palette.Muted, Font = new Font(Font.FontFamily, 8f, FontStyle.Bold) };
        var row = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, BackColor = _palette.Card };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 52));
        _operator.Dock = DockStyle.Fill;
        _operator.DropDownStyle = ComboBoxStyle.DropDownList;
        _operator.FlatStyle = FlatStyle.Flat;
        _operator.BackColor = _palette.Input;
        _operator.ForeColor = _palette.Text;
        _operator.Items.AddRange(new object[] { "No condition", "Contains", "Does not contain", "Equals", "Does not equal", "Starts with", "Ends with", "Is blank", "Is not blank" });
        _operator.SelectedIndex = current == null ? 0 : (int)current.Operator;
        _operator.SelectedIndexChanged += (_, _) => _operand.Enabled = _operator.SelectedIndex is >= 1 and <= 6;
        _operand.Dock = DockStyle.Fill;
        _operand.Margin = new Padding(8, 0, 0, 0);
        _operand.BorderStyle = BorderStyle.FixedSingle;
        _operand.BackColor = _palette.Input;
        _operand.ForeColor = _palette.Text;
        _operand.Text = current?.Operand ?? string.Empty;
        _operand.Enabled = _operator.SelectedIndex is >= 1 and <= 6;
        row.Controls.Add(_operator, 0, 0);
        row.Controls.Add(_operand, 1, 0);
        card.Controls.Add(row);
        card.Controls.Add(label);
        return card;
    }

    private Control BuildSearch()
    {
        var panel = new Panel { Dock = DockStyle.Top, Height = 62, Padding = new Padding(18, 15, 18, 10), BackColor = _palette.Surface };
        _search.Dock = DockStyle.Fill;
        _search.BorderStyle = BorderStyle.FixedSingle;
        _search.BackColor = _palette.Input;
        _search.ForeColor = _palette.Text;
        _search.TextChanged += (_, _) => ApplySearch();
        panel.Controls.Add(_search);
        return panel;
    }

    private Control BuildValueToolbar()
    {
        var panel = new Panel { Dock = DockStyle.Top, Height = 42, Padding = new Padding(18, 5, 18, 5), BackColor = _palette.Surface };
        _selectionLabel.AutoSize = true;
        _selectionLabel.ForeColor = _palette.Muted;
        _selectionLabel.Location = new Point(18, 11);

        var clear = MakeLink("Clear");
        clear.Location = new Point(0, 7);
        clear.Click += (_, _) => SetVisible(false);
        var all = MakeLink("All");
        all.Location = new Point(0, 7);
        all.Click += (_, _) => SetVisible(true);
        panel.Controls.Add(_selectionLabel);
        panel.Controls.Add(all);
        panel.Controls.Add(clear);
        panel.Resize += (_, _) =>
        {
            clear.Location = new Point(panel.ClientSize.Width - clear.Width - 18, 7);
            all.Location = new Point(clear.Left - all.Width - 4, 7);
        };
        UpdateSelectionLabel();
        return panel;
    }

    private Control BuildValues()
    {
        var host = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 0, 12, 8), BackColor = _palette.Surface };
        _valueList.Dock = DockStyle.Fill;
        _valueList.FlowDirection = FlowDirection.TopDown;
        _valueList.WrapContents = false;
        _valueList.AutoScroll = true;
        _valueList.BackColor = _palette.Input;
        _valueList.Padding = new Padding(8, 7, 8, 7);
        foreach (var value in _allValues)
        {
            var check = new CheckBox
            {
                Text = string.IsNullOrEmpty(value) ? "(blank)" : value,
                Tag = value,
                Checked = _selectedValues.Contains(value),
                AutoSize = false,
                Width = 370,
                Height = 28,
                ForeColor = _palette.Text,
                BackColor = _palette.Input,
                FlatStyle = FlatStyle.Flat,
                Padding = new Padding(3, 0, 0, 0)
            };
            check.CheckedChanged += (_, _) =>
            {
                if (check.Checked) _selectedValues.Add((string)check.Tag);
                else _selectedValues.Remove((string)check.Tag);
                UpdateSelectionLabel();
            };
            _checks[value] = check;
            _valueList.Controls.Add(check);
        }
        host.Controls.Add(_valueList);
        return host;
    }

    private Control BuildFooter()
    {
        var panel = new Panel { Dock = DockStyle.Bottom, Height = 68, Padding = new Padding(18, 14, 18, 14), BackColor = _palette.Card };
        var clearFilter = MakeButton("Clear filter", _palette.Card, _palette.Text, 94, 38);
        clearFilter.Location = new Point(18, 15);
        clearFilter.Click += (_, _) => { Result = null; DialogResult = DialogResult.OK; Close(); };

        var cancel = MakeButton("Cancel", _palette.Card, _palette.Text, 76, 38);
        cancel.Location = new Point(0, 15);
        cancel.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
        var apply = MakeButton("Apply", _palette.Accent, Color.White, 82, 38);
        _applyButton = apply;
        apply.Enabled = true;
        apply.TabStop = true;
        apply.DialogResult = DialogResult.OK;
        apply.Location = new Point(0, 15);
        apply.Font = new Font(Font.FontFamily, 9.5f, FontStyle.Bold);
        apply.Click += (_, _) => CommitFilter();
        panel.Controls.Add(clearFilter);
        panel.Controls.Add(cancel);
        panel.Controls.Add(apply);
        panel.Resize += (_, _) =>
        {
            apply.Location = new Point(panel.ClientSize.Width - apply.Width - 18, 15);
            cancel.Location = new Point(apply.Left - cancel.Width - 8, 15);
        };
        AcceptButton = apply;
        CancelButton = cancel;
        return panel;
    }

    private Button MakeButton(string text, Color back, Color fore, int width, int height) => new()
    {
        Text = text,
        Width = width,
        Height = height,
        BackColor = back,
        ForeColor = fore,
        FlatStyle = FlatStyle.Flat,
        Cursor = Cursors.Hand,
        UseVisualStyleBackColor = false,
        FlatAppearance = { BorderColor = _palette.Border, BorderSize = 1, MouseOverBackColor = _palette.Hover }
    };

    private Button MakeLink(string text)
    {
        var button = MakeButton(text, _palette.Surface, _palette.Accent, 44, 28);
        button.FlatAppearance.BorderSize = 0;
        return button;
    }

    private void ApplySearch()
    {
        var search = _search.Text;
        foreach (var pair in _checks)
            pair.Value.Visible = pair.Key.IndexOf(search, StringComparison.CurrentCultureIgnoreCase) >= 0;
    }

    private void SetVisible(bool value)
    {
        foreach (var check in _checks.Values.Where(c => c.Visible)) check.Checked = value;
        UpdateSelectionLabel();
    }

    private void UpdateSelectionLabel()
    {
        if (_selectionLabel.IsDisposed) return;
        _selectionLabel.Text = _truncated
            ? $"{_selectedValues.Count:n0} selected · first 1,000 values"
            : $"{_selectedValues.Count:n0} of {_allValues.Count:n0} selected";
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Enter)
        {
            CommitFilter();
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (DialogResult == DialogResult.OK && Result == null) BuildResult();
        base.OnFormClosing(e);
    }

    private void BuildResult()
    {
        Result = new FilterSpec
        {
            AllowedValues = _selectedValues.Count == _allValues.Count ? null : new HashSet<string>(_selectedValues, StringComparer.CurrentCultureIgnoreCase),
            Operator = (TextFilterOperator)_operator.SelectedIndex,
            Operand = _operand.Text
        };
    }

    private void CommitFilter()
    {
        BuildResult();
        DialogResult = DialogResult.OK;
        Close();
    }

    private void ApplyRoundedRegion()
    {
        using var path = new System.Drawing.Drawing2D.GraphicsPath();
        const int radius = 12;
        path.AddArc(0, 0, radius, radius, 180, 90);
        path.AddArc(Width - radius, 0, radius, radius, 270, 90);
        path.AddArc(Width - radius, Height - radius, radius, radius, 0, 90);
        path.AddArc(0, Height - radius, radius, radius, 90, 90);
        path.CloseFigure();
        Region = new Region(path);
    }

    private void DragWindow(object sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left) return;
        ReleaseCapture();
        SendMessage(Handle, 0xA1, (IntPtr)2, IntPtr.Zero);
    }

    private static void SetCueBanner(TextBox box, string text) => SendMessage(box.Handle, 0x1501, (IntPtr)1, text);

    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, string lParam);

    private readonly struct Palette
    {
        public Palette(Color surface, Color card, Color input, Color text, Color muted, Color border, Color hover, Color accent)
        {
            Surface = surface; Card = card; Input = input; Text = text; Muted = muted; Border = border; Hover = hover; Accent = accent;
        }
        public Color Surface { get; }
        public Color Card { get; }
        public Color Input { get; }
        public Color Text { get; }
        public Color Muted { get; }
        public Color Border { get; }
        public Color Hover { get; }
        public Color Accent { get; }
        public static Palette Dark => new(Color.FromArgb(31, 31, 31), Color.FromArgb(38, 38, 38), Color.FromArgb(45, 45, 48), Color.FromArgb(242, 242, 242), Color.FromArgb(166, 166, 166), Color.FromArgb(68, 68, 72), Color.FromArgb(57, 57, 61), Color.FromArgb(0, 122, 204));
        public static Palette Light => new(Color.FromArgb(250, 250, 250), Color.White, Color.White, Color.FromArgb(32, 32, 32), Color.FromArgb(100, 100, 100), Color.FromArgb(210, 210, 210), Color.FromArgb(238, 238, 238), Color.FromArgb(0, 120, 212));
    }
}
