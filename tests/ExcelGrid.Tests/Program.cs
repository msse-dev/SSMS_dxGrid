using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using ExcelGrid.Ssms;
using Microsoft.SqlServer.Management.UI.Grid;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        try
        {
            SnapshotReadsNativeStorageWithoutLegacyProxy();
            OnlyQueryResultGridsAreEligibleForReplacement();
            DevExpressGridActuallyFiltersAndSortsRows();
            SelectAllCornerSelectsVisibleCells();
            NativeSelectionCompressionPreservesTabularCopy();
            DevExpressSurfaceIsHostedInsideNativeGrid();
            NativeContextMenuRelayCanRepeat();
            Console.WriteLine("All active ExcelGrid tests passed.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static void SnapshotReadsNativeStorageWithoutLegacyProxy()
    {
        using var native = CreateNativeGrid(
            new object[] { "Heritage Blue" },
            new object[] { "Heritage Choc" });

        var snapshot = QueryResultSnapshot.Capture(native);
        Equal(2L, snapshot.SourceRowCount);
        Equal(1, snapshot.Table.Columns.Count);
        Equal("cBreed", snapshot.Captions[0]);
        Equal(typeof(string), snapshot.Table.Columns[0].DataType);
        Equal("Heritage Choc", snapshot.Table.Rows[1][0]);
    }

    private static void OnlyQueryResultGridsAreEligibleForReplacement()
    {
        using var administrativeGrid = CreateNativeGrid(new object[] { "Job step" });
        Equal(false, GridDiscovery.IsQueryResultsGrid(administrativeGrid));

        using var queryGrid = new Microsoft.SqlServer.Management.UI.VSIntegration.Editors.GridResultsGrid();
        Equal(true, GridDiscovery.IsQueryResultsGrid(queryGrid));
    }

    private static void DevExpressGridActuallyFiltersAndSortsRows()
    {
        var table = new DataTable();
        table.Columns.Add("C0", typeof(string));
        table.Rows.Add("Heirloom");
        table.Rows.Add("Heritage Blue");
        table.Rows.Add("Heritage Choc");
        table.Rows.Add("Heritage Mixed");

        using var grid = new DevExpressResultsControl(true);
        grid.CreateControl();
        Equal(true, grid.IsDarkTheme);
        Equal("DevExpress Dark Style", grid.SkinName);
        Equal(Color.FromArgb(37, 37, 38), grid.RowBackColor);
        Equal(Color.FromArgb(51, 51, 55), grid.HeaderBackColor);
        Equal(Color.FromArgb(78, 78, 84), grid.GridLineColor);
        grid.ApplyTheme(false);
        Equal(false, grid.IsDarkTheme);
        Equal("Office 2019 Colorful", grid.SkinName);
        Equal(true, SsmsTheme.IsDark(Color.FromArgb(37, 37, 38)));
        Equal(false, SsmsTheme.IsDark(Color.FromArgb(245, 245, 245)));

        grid.SetData(table, new[] { "cBreed" }, 4, false);
        Equal(4, grid.View.DataRowCount);
        Equal(DevExpress.XtraGrid.Columns.AutoFilterCondition.Contains, grid.View.Columns[0].OptionsFilter.AutoFilterCondition);
        Equal(DevExpress.XtraGrid.Columns.ExcelFilterDefaultTab.Filters, grid.View.Columns[0].OptionsFilter.PopupExcelFilterDefaultTab);
        Equal(DevExpress.XtraGrid.Columns.ExcelFilterTextFilters.AllFilters, grid.View.Columns[0].OptionsFilter.PopupExcelFilterTextFilters);
        Equal(DevExpress.Utils.DefaultBoolean.False, grid.View.OptionsFilter.AllowAutoFilterConditionChange);

        grid.View.SetRowCellValue(DevExpress.XtraGrid.GridControl.AutoFilterRowHandle, grid.View.Columns[0], "Choc");
        Equal(true, grid.View.ActiveFilterString.IndexOf("Contains", StringComparison.OrdinalIgnoreCase) >= 0);
        Equal(1, grid.View.DataRowCount);
        Equal("Heritage Choc", (string)grid.View.GetRowCellValue(0, "C0"));

        grid.View.ClearColumnsFilter();
        grid.View.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
        var visibleRow = grid.View.GetVisibleRowHandle(0);
        Equal(true, grid.TryMapNativeCell(visibleRow, grid.View.Columns[0], out var nativeRow, out var nativeColumn));
        Equal(3L, nativeRow);
        Equal(1, nativeColumn);
    }

    private static void DevExpressSurfaceIsHostedInsideNativeGrid()
    {
        using var native = CreateNativeGrid(new object[] { "Heritage Choc" });
        native.BackColor = Color.White;
        using var enhancer = new DevExpressGridEnhancer(native);
        var replacement = native.Controls.OfType<DevExpressResultsControl>().Single();

        Equal(false, replacement.IsDarkTheme);
        Equal(1, replacement.View.DataRowCount);
        native.BackColor = Color.FromArgb(37, 37, 38);
        enhancer.Refresh();
        Equal(true, replacement.IsDarkTheme);
        Equal(true, ReferenceEquals(native, replacement.Parent));
        Equal(System.Windows.Forms.DockStyle.Fill, replacement.Dock);
        Equal(true, replacement.Visible);
        Equal(true, native.Visible);
    }

    private static void SelectAllCornerSelectsVisibleCells()
    {
        var table = new DataTable();
        table.Columns.Add("C0", typeof(string));
        table.Columns.Add("C1", typeof(int));
        table.Rows.Add("Heirloom", 30);
        table.Rows.Add("Heritage Blue", 8);
        table.Rows.Add("Heritage Choc", 9);
        table.Rows.Add("Heritage Mixed", 14);

        using var grid = new DevExpressResultsControl(true);
        grid.Size = new Size(800, 300);
        grid.CreateControl();
        grid.SetData(table, new[] { "cBreed", "iBreed_PK" }, 4, false);
        grid.PerformLayout();
        System.Windows.Forms.Application.DoEvents();
        Equal(true, grid.TrySelectAllFromCorner(new Point(5, 5)));
        Equal(8, grid.View.GetSelectedCells().Length);

        grid.View.ClearSelection();
        grid.View.SetRowCellValue(DevExpress.XtraGrid.GridControl.AutoFilterRowHandle, grid.View.Columns[0], "Choc");
        Equal(1, grid.View.DataRowCount);
        Equal(true, grid.TrySelectAllFromCorner(new Point(5, 5)));
        Equal(2, grid.View.GetSelectedCells().Length);
        Equal("Heritage Choc", (string)grid.View.GetRowCellValue(grid.View.GetSelectedCells()[0].RowHandle, "C0"));
    }

    private static void NativeContextMenuRelayCanRepeat()
    {
        using var native = CreateNativeGrid(new object[] { "Heritage Choc" });
        var nativeRightClickCount = 0;
        var nativeMouseDownCount = 0;
        var nativeFocusedDuringRightClick = false;

        native.MouseDown += (_, args) =>
        {
            if (args.Button == System.Windows.Forms.MouseButtons.Right) nativeMouseDownCount++;
        };
        native.MouseButtonClicked += (_, args) =>
        {
            if (args.RowIndex == 0 && args.ColumnIndex == 1 && args.Button == System.Windows.Forms.MouseButtons.Right)
            {
                nativeRightClickCount++;
                nativeFocusedDuringRightClick |= native.Focused;
            }
        };

        using (var enhancer = new DevExpressGridEnhancer(native))
        {
            for (var attempt = 0; attempt < 5; attempt++)
            {
                enhancer.RelayNativeContextMenu(0, 1);
                Equal(false, native.Capture);
                Equal(true, native.Focused);
            }
        }

        Equal(5, nativeRightClickCount);
        Equal(5, nativeMouseDownCount);
        Equal(true, nativeFocusedDuringRightClick);
    }

    private static void NativeSelectionCompressionPreservesTabularCopy()
    {
        var cells = new List<NativeCell>();
        for (long row = 0; row < 1_099; row++)
            for (var column = 1; column <= 12; column++)
                cells.Add(new NativeCell(row, column));

        var compressed = DevExpressGridEnhancer.BuildNativeSelection(cells);
        Equal(1, compressed.Count);
        Equal(0L, compressed[0].Y);
        Equal(1, compressed[0].X);
        Equal(1_099L, compressed[0].Height);
        Equal(12, compressed[0].Width);

        using var native = CreateNativeGrid(
            new object[] { "Heritage Blue", 8 },
            new object[] { "Heritage Choc", 9 });
        var selection = DevExpressGridEnhancer.BuildNativeSelection(new[]
        {
            new NativeCell(0, 1), new NativeCell(0, 2),
            new NativeCell(1, 1), new NativeCell(1, 2)
        });
        native.SelectionType = GridSelectionType.CellBlocks;
        native.SelectedCells = selection;

        var copyMethod = typeof(GridControl).GetMethod("GetClipboardTextForSelectionBlock",
            BindingFlags.Instance | BindingFlags.NonPublic)!;
        var withoutHeaders = (string)copyMethod.Invoke(native, new object[] { 0, false })!;
        var withHeaders = (string)copyMethod.Invoke(native, new object[] { 0, true })!;
        Equal(true, withoutHeaders.Contains("Heritage Blue\t8"));
        Equal(true, withoutHeaders.Contains("\r\nHeritage Choc\t9"));
        Equal(true, withHeaders.StartsWith("cBreed\tiBreed_PK\r\n", StringComparison.Ordinal));
    }

    private static GridControl CreateNativeGrid(params object[][] rows)
    {
        var native = new GridControl { Size = new Size(800, 300), Visible = true };
        native.AddColumn(new GridColumnInfo()); // SSMS row-number margin
        var columnCount = rows.Length == 0 ? 1 : rows[0].Length;
        for (var column = 0; column < columnCount; column++)
        {
            native.AddColumn(new GridColumnInfo());
            native.SetHeaderInfo(column + 1, column == 0 ? "cBreed" : "iBreed_PK", (Bitmap)null!);
        }
        native.GridStorage = new FakeGridStorage(rows);
        native.UpdateGrid(true);
        native.CreateControl();
        return native;
    }

    private static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException($"Expected '{expected}', got '{actual}'.");
    }
}

internal sealed class FakeGridStorage : IGridStorage
{
    // Matches QEResultSet's active adapter shape used by QueryResultSnapshot.
    private readonly FakeStorageView m_view;

    public FakeGridStorage(params object[][] rows) => m_view = new FakeStorageView(rows);
    public long NumRows() => m_view.NumRows();
    public long EnsureRowsInBuf(long first, long last) => m_view.EnsureRowsInBuf(first, last);
    public Type GetFieldType(int column) => m_view.NumRows() == 0 || m_view.GetCellData(0, column) == null
        ? typeof(object)
        : m_view.GetCellData(0, column).GetType();
    // GridControl passes its UI/storage column number (one-based after the row
    // margin); QueryResultSnapshot reads the zero-based storage view directly.
    public string GetCellDataAsString(long row, int column) => m_view.GetCellDataAsString(row, column - 1);
    public int IsCellEditable(long row, int column) => 0;
    public Bitmap GetCellDataAsBitmap(long row, int column) => null!;
    public void GetCellDataForButton(long row, int column, out ButtonCellState state, out Bitmap bitmap, out string text)
    {
        state = default;
        bitmap = null!;
        text = GetCellDataAsString(row, column);
    }
    public GridCheckBoxState GetCellDataForCheckBox(long row, int column) => default;
    public void FillControlWithData(long row, int column, IGridEmbeddedControl control) { }
    public bool SetCellDataFromControl(long row, int column, IGridEmbeddedControl control) => false;
}

internal sealed class FakeStorageView : IStorageView
{
    private readonly object[][] _rows;

    public FakeStorageView(params object[][] rows) => _rows = rows;
    public long NumRows() => _rows.LongLength;
    public long EnsureRowsInBuf(long first, long last) => Math.Max(0, last - first + 1);
    public string GetCellDataAsString(long row, int column) => Convert.ToString(GetCellData(row, column)) ?? string.Empty;
    public object GetCellData(long row, int column) => _rows[row][column];
    public void DeleteRow(long row) => throw new NotSupportedException();
    public IColumnInfo GetColumnInfo(int column) => null!;
    public int NumColumns() => _rows.Length == 0 ? 0 : _rows[0].Length;
    public bool IsStorageClosed() => false;
    public void Dispose() { }
}

namespace Microsoft.SqlServer.Management.UI.VSIntegration.Editors
{
    // Matches the private SSMS query-results control name without taking a build
    // dependency on SQLEditors.dll. Administrative grids do not use this type.
    internal sealed class GridResultsGrid : GridControl { }
}
