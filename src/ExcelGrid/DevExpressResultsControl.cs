using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace ExcelGrid.Ssms;

internal sealed class DevExpressResultsControl : UserControl
{
    private static readonly string VersionLabel = "Excel Grid v" +
        typeof(DevExpressResultsControl).Assembly.GetName().Version.ToString(3);

    private readonly GridControl _grid = new();
    private readonly GridView _view;
    private readonly LabelControl _status = new();
    private readonly SimpleButton _clear = new();
    private readonly SimpleButton _bestFit = new();
    private readonly PanelControl _toolbar;
    private bool _dark;
    private bool _themeApplied;

    public DevExpressResultsControl(bool dark)
    {
        Dock = DockStyle.Fill;
        _view = new GridView(_grid);
        _grid.MainView = _view;
        _grid.ViewCollection.Add(_view);
        _grid.Dock = DockStyle.Fill;

        _view.OptionsBehavior.Editable = false;
        _view.OptionsBehavior.AllowIncrementalSearch = true;
        _view.OptionsView.ShowAutoFilterRow = true;
        _view.OptionsView.ShowGroupPanel = false;
        _view.OptionsView.ShowFooter = true;
        _view.OptionsView.ColumnAutoWidth = false;
        _view.OptionsSelection.MultiSelect = true;
        _view.OptionsSelection.MultiSelectMode = GridMultiSelectMode.CellSelect;
        _view.OptionsClipboard.CopyColumnHeaders = DevExpress.Utils.DefaultBoolean.True;
        _view.OptionsFilter.ColumnFilterPopupMode = DevExpress.XtraGrid.Columns.ColumnFilterPopupMode.Excel;
        _view.OptionsFilter.ShowAllTableValuesInCheckedFilterPopup = true;
        _view.OptionsFilter.AllowMultiSelectInCheckedFilterPopup = true;
        _view.OptionsFilter.AllowAutoFilterConditionChange = DevExpress.Utils.DefaultBoolean.False;
        _view.ColumnFilterChanged += (_, _) => UpdateStatus();
        _view.PopupMenuShowing += (_, e) => e.Allow = false;
        _view.RowCellStyle += (_, e) =>
        {
            if (_dark && e.RowHandle == GridControl.AutoFilterRowHandle)
                SetAppearance(e.Appearance, Color.FromArgb(30, 30, 30), Color.FromArgb(250, 250, 250));
        };
        _grid.MouseDown += GridMouseDown;

        _toolbar = new PanelControl
        {
            Dock = DockStyle.Top,
            Height = 34,
            BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        };
        _toolbar.LookAndFeel.UseDefaultLookAndFeel = false;
        _toolbar.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
        _toolbar.Appearance.Options.UseBackColor = true;

        _status.Location = new Point(10, 9);
        _status.AutoSizeMode = LabelAutoSizeMode.None;
        _status.Size = new Size(620, 22);
        _status.Text = VersionLabel + " | Auto-filter row and header dropdowns offer Contains";
        _status.Appearance.BackColor = Color.Transparent;
        _status.Appearance.Options.UseForeColor = true;
        _status.Appearance.Options.UseBackColor = true;

        _clear.Text = "Clear filters";
        _clear.Size = new Size(88, 26);
        _clear.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        _clear.Click += (_, _) => { _view.ClearColumnsFilter(); _view.ClearSorting(); };
        _bestFit.Text = "Best fit";
        _bestFit.Size = new Size(72, 26);
        _bestFit.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        _bestFit.Click += (_, _) => _view.BestFitColumns();
        _toolbar.Controls.Add(_status);
        _toolbar.Controls.Add(_clear);
        _toolbar.Controls.Add(_bestFit);
        _toolbar.Resize += (_, _) =>
        {
            _clear.Location = new Point(_toolbar.ClientSize.Width - _clear.Width - 8, 4);
            _bestFit.Location = new Point(_clear.Left - _bestFit.Width - 6, 4);
        };

        Controls.Add(_grid);
        Controls.Add(_toolbar);
        ApplyTheme(dark);
    }

    internal GridView View => _view;
    internal bool IsDarkTheme => _dark;
    internal string SkinName => _grid.LookAndFeel.SkinName;
    internal Color RowBackColor => _view.Appearance.Row.BackColor;
    internal Color HeaderBackColor => _view.Appearance.HeaderPanel.BackColor;
    internal Color GridLineColor => _view.Appearance.HorzLine.BackColor;
    internal event EventHandler<NativeContextMenuRequestEventArgs>? NativeContextMenuRequested;

    public void ApplyTheme(bool dark)
    {
        if (_themeApplied && _dark == dark) return;

        _dark = dark;
        _themeApplied = true;
        var skin = dark ? "DevExpress Dark Style" : "Office 2019 Colorful";
        ApplySkin(_grid.LookAndFeel, skin);
        ApplySkin(_toolbar.LookAndFeel, skin);
        ApplySkin(_status.LookAndFeel, skin);
        ApplySkin(_clear.LookAndFeel, skin);
        ApplySkin(_bestFit.LookAndFeel, skin);

        BackColor = dark ? Color.FromArgb(30, 30, 30) : Color.White;
        ForeColor = dark ? Color.FromArgb(241, 241, 241) : Color.FromArgb(32, 32, 32);
        _toolbar.Appearance.BackColor = dark ? Color.FromArgb(37, 37, 38) : Color.FromArgb(245, 245, 245);
        _status.Appearance.ForeColor = ForeColor;
        ApplyGridPalette(dark);
        _view.LayoutChanged();
        _grid.Invalidate(true);
        _toolbar.Invalidate(true);
    }

    public void SetData(DataTable table, IReadOnlyList<string> captions, long sourceRows, bool truncated)
    {
        _grid.DataSource = null;
        _grid.DataSource = table;
        _view.PopulateColumns();
        for (var index = 0; index < _view.Columns.Count && index < captions.Count; index++)
        {
            var column = _view.Columns[index];
            column.Caption = captions[index];
            column.OptionsFilter.AllowFilter = true;
            column.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.True;
            column.OptionsFilter.ImmediateUpdatePopupExcelFilter = DevExpress.Utils.DefaultBoolean.True;
            column.OptionsFilter.PopupExcelFilterDefaultTab = DevExpress.XtraGrid.Columns.ExcelFilterDefaultTab.Filters;
            column.OptionsFilter.PopupExcelFilterTextFilters = DevExpress.XtraGrid.Columns.ExcelFilterTextFilters.AllFilters;
            column.OptionsFilter.AutoFilterCondition = DevExpress.XtraGrid.Columns.AutoFilterCondition.Contains;
            column.OptionsFilter.ImmediateUpdateAutoFilter = true;
            column.MinWidth = 90;
        }
        if (_view.Columns.Count > 0)
            _view.Columns[0].Summary.Add(DevExpress.Data.SummaryItemType.Count, _view.Columns[0].FieldName, "{0:n0} rows");
        _view.BestFitMaxRowCount = 200;
        _view.BestFitColumns();
        Tag = new ResultInfo(sourceRows, truncated);
        _bestFit.Enabled = true;
        _clear.Enabled = true;
        UpdateStatus();
    }

    public void ShowError(string message)
    {
        _grid.DataSource = null;
        Tag = null;
        _status.Text = VersionLabel + " | " + message;
    }

    public void ShowLoading(string message)
    {
        _grid.DataSource = null;
        Tag = null;
        _status.Text = $"{VersionLabel} | {message}";
        _status.Appearance.ForeColor = _dark ? Color.White : Color.FromArgb(32, 32, 32);
        _bestFit.Enabled = false;
        _clear.Enabled = false;
    }

    private void UpdateStatus()
    {
        var info = Tag as ResultInfo;
        var visible = _view.DataRowCount;
        var total = info?.SourceRows ?? visible;
        var suffix = info?.Truncated == true ? " (first 250,000 loaded)" : string.Empty;
        _status.Text = $"{VersionLabel} | {visible:n0} of {total:n0} rows{suffix} | Contains is the default text filter";
    }

    private void GridMouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left && TrySelectAllFromCorner(e.Location)) return;

        if (e.Button != MouseButtons.Right) return;
        var hit = _view.CalcHitInfo(e.Location);
        var column = hit.Column;
        if (column == null || hit.RowHandle < 0) return;

        if (!_view.IsCellSelected(hit.RowHandle, column))
        {
            _view.ClearSelection();
            _view.SelectCell(hit.RowHandle, column);
        }
        _view.FocusedRowHandle = hit.RowHandle;
        _view.FocusedColumn = column;
        if (TryMapNativeCell(hit.RowHandle, column, out var nativeRow, out var nativeColumn))
        {
            var selectedCells = new List<NativeCell>();
            foreach (var cell in _view.GetSelectedCells()
                         .OrderBy(cell => _view.GetVisibleIndex(cell.RowHandle))
                         .ThenBy(cell => cell.Column.VisibleIndex))
                if (TryMapNativeCell(cell.RowHandle, cell.Column, out var selectedRow, out var selectedColumn))
                    selectedCells.Add(new NativeCell(selectedRow, selectedColumn));
            NativeContextMenuRequested?.Invoke(this,
                new NativeContextMenuRequestEventArgs(nativeRow, nativeColumn, selectedCells, _grid.PointToScreen(e.Location)));
        }
    }

    private static bool IsSelectAllCorner(GridHitInfo hit) =>
        (hit.HitTest == GridHitTest.ColumnButton && hit.Column == null) ||
        (hit.HitTest == GridHitTest.RowIndicator && hit.RowHandle == GridControl.AutoFilterRowHandle);

    internal bool TrySelectAllFromCorner(Point location)
    {
        if (!IsSelectAllCorner(_view.CalcHitInfo(location))) return false;
        SelectAllVisibleCells();
        return true;
    }

    internal void SelectAllVisibleCells()
    {
        _view.SelectAll();
        _view.Invalidate();
    }

    internal bool TryMapNativeCell(int rowHandle, GridColumn column, out long nativeRow, out int nativeColumn)
    {
        nativeRow = -1;
        nativeColumn = -1;
        if (rowHandle < 0 || column == null) return false;

        var sourceIndex = _view.GetDataSourceRowIndex(rowHandle);
        if (sourceIndex < 0) return false;
        nativeRow = sourceIndex;
        nativeColumn = column.AbsoluteIndex + 1; // SSMS column zero is its row-number margin.

        if (_grid.DataSource is DataTable table &&
            table.Columns.Contains(column.FieldName) &&
            table.Columns[column.FieldName].ExtendedProperties["SsmsUiColumnIndex"] is int uiColumn)
            nativeColumn = uiColumn;
        return true;
    }

    private static void ApplySkin(DevExpress.LookAndFeel.UserLookAndFeel lookAndFeel, string skin)
    {
        lookAndFeel.UseDefaultLookAndFeel = false;
        lookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Skin;
        lookAndFeel.SkinName = skin;
    }

    private void ApplyGridPalette(bool dark)
    {
        var appearances = new[]
        {
            _view.Appearance.Row, _view.Appearance.EvenRow, _view.Appearance.OddRow,
            _view.Appearance.HeaderPanel, _view.Appearance.Empty, _view.Appearance.FocusedCell,
            _view.Appearance.FocusedRow, _view.Appearance.SelectedRow, _view.Appearance.HideSelectionRow,
            _view.Appearance.HorzLine, _view.Appearance.VertLine, _view.Appearance.RowSeparator,
            _view.Appearance.FooterPanel, _view.Appearance.FilterPanel,
            _view.Appearance.ColumnFilterButton, _view.Appearance.ColumnFilterButtonActive
        };
        foreach (var appearance in appearances) appearance.Reset();
        _clear.Appearance.Reset();
        _clear.AppearanceHovered.Reset();
        _clear.AppearancePressed.Reset();
        _bestFit.Appearance.Reset();
        _bestFit.AppearanceHovered.Reset();
        _bestFit.AppearancePressed.Reset();

        _view.OptionsView.EnableAppearanceEvenRow = dark;
        _view.OptionsView.EnableAppearanceOddRow = dark;
        _grid.BackColor = dark ? Color.FromArgb(30, 30, 30) : Color.White;
        if (!dark) return;

        var text = Color.FromArgb(245, 245, 245);
        var secondaryText = Color.FromArgb(225, 225, 225);
        var row = Color.FromArgb(37, 37, 38);
        var alternateRow = Color.FromArgb(43, 43, 46);
        var header = Color.FromArgb(51, 51, 55);
        var gridLine = Color.FromArgb(78, 78, 84);
        var selection = Color.FromArgb(9, 71, 113);

        SetAppearance(_view.Appearance.Row, row, text);
        SetAppearance(_view.Appearance.OddRow, row, text);
        SetAppearance(_view.Appearance.EvenRow, alternateRow, text);
        SetAppearance(_view.Appearance.HeaderPanel, header, Color.White);
        SetAppearance(_view.Appearance.Empty, Color.FromArgb(30, 30, 30), text);
        SetAppearance(_view.Appearance.FocusedCell, selection, Color.White);
        SetAppearance(_view.Appearance.FocusedRow, row, text);
        SetAppearance(_view.Appearance.SelectedRow, selection, Color.White);
        SetAppearance(_view.Appearance.HideSelectionRow, Color.FromArgb(62, 62, 68), Color.White);
        SetAppearance(_view.Appearance.HorzLine, gridLine, gridLine);
        SetAppearance(_view.Appearance.VertLine, gridLine, gridLine);
        SetAppearance(_view.Appearance.RowSeparator, gridLine, gridLine);
        SetAppearance(_view.Appearance.FooterPanel, header, secondaryText);
        SetAppearance(_view.Appearance.FilterPanel, row, secondaryText);
        SetAppearance(_view.Appearance.ColumnFilterButton, header, secondaryText);
        SetAppearance(_view.Appearance.ColumnFilterButtonActive, selection, Color.White);

        ApplyButtonPalette(_clear);
        ApplyButtonPalette(_bestFit);
    }

    private static void ApplyButtonPalette(SimpleButton button)
    {
        SetAppearance(button.Appearance, Color.FromArgb(62, 62, 66), Color.White);
        SetAppearance(button.AppearanceHovered, Color.FromArgb(80, 80, 85), Color.White);
        SetAppearance(button.AppearancePressed, Color.FromArgb(0, 122, 204), Color.White);
    }

    private static void SetAppearance(DevExpress.Utils.AppearanceObject appearance, Color backColor, Color foreColor)
    {
        appearance.BackColor = backColor;
        appearance.ForeColor = foreColor;
        appearance.Options.UseBackColor = true;
        appearance.Options.UseForeColor = true;
    }

    private sealed class ResultInfo
    {
        public ResultInfo(long sourceRows, bool truncated) { SourceRows = sourceRows; Truncated = truncated; }
        public long SourceRows { get; }
        public bool Truncated { get; }
    }
}

internal sealed class NativeContextMenuRequestEventArgs : EventArgs
{
    public NativeContextMenuRequestEventArgs(long row, int column, IReadOnlyList<NativeCell> selectedCells, Point screenLocation)
    {
        Row = row;
        Column = column;
        SelectedCells = selectedCells;
        ScreenLocation = screenLocation;
    }

    public long Row { get; }
    public int Column { get; }
    public IReadOnlyList<NativeCell> SelectedCells { get; }
    public Point ScreenLocation { get; }
}

internal readonly struct NativeCell
{
    public NativeCell(long row, int column) { Row = row; Column = column; }
    public long Row { get; }
    public int Column { get; }
}
