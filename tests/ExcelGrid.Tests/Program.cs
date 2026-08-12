using System;
using System.Collections.Generic;
using System.Drawing;
using System.Data;
using System.IO;
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
            NumericSortCyclesAndRemainsStable();
            FiltersCombineAndProxyPreservesRuntimeInterface();
            DialogReturnsUncheckedValues();
            DevExpressGridActuallyFiltersRows();
            DevExpressSurfaceIsHostedInsideNativeGrid();
            ProxyFitsTheRealSsmsResultSetField();
            Console.WriteLine("All ExcelGrid storage tests passed.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static void NumericSortCyclesAndRemainsStable()
    {
        var storage = new SortableFilterStorage(new FakeStorage(
            new[] { "ten", "2" }, new[] { "two-a", "10" }, new[] { "two-b", "2" }));

        storage.CycleSort(1);
        Equal("ten", storage.GetCellDataAsString(0, 0));
        Equal("two-b", storage.GetCellDataAsString(1, 0));
        Equal("two-a", storage.GetCellDataAsString(2, 0));

        storage.CycleSort(1);
        Equal("two-a", storage.GetCellDataAsString(0, 0));
        Equal("ten", storage.GetCellDataAsString(1, 0));
        Equal("two-b", storage.GetCellDataAsString(2, 0));

        storage.CycleSort(1);
        Equal("ten", storage.GetCellDataAsString(0, 0));
    }

    private static void FiltersCombineAndProxyPreservesRuntimeInterface()
    {
        var inner = new FakeStorage(
            new[] { "Alice", "Active" }, new[] { "Bob", "Inactive" }, new[] { "Alicia", "Active" });
        var storage = new SortableFilterStorage(inner);
        storage.SetFilter(0, new FilterSpec { Operator = TextFilterOperator.StartsWith, Operand = "Ali" });
        storage.SetFilter(1, new FilterSpec { AllowedValues = new HashSet<string> { "Active" } });
        Equal(2L, storage.NumRows());
        Equal("Alice", storage.GetCellDataAsString(0, 0));
        Equal("Alicia", storage.GetCellDataAsString(1, 0));

        var proxy = (ITestStorageView)StorageViewRealProxy.Create(typeof(ITestStorageView), storage);
        Equal(2L, proxy.NumRows());
        Equal("Alicia", proxy.GetCellDataAsString(1, 0));
        proxy.MaxNumBytesToDisplay = 512;
        Equal(512, inner.MaxNumBytesToDisplay);
    }

    private static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException($"Expected '{expected}', got '{actual}'.");
    }

    private static void ProxyFitsTheRealSsmsResultSetField()
    {
        var ssmsRoot = Environment.GetEnvironmentVariable("EXCELGRID_SSMS_ROOT");
        if (string.IsNullOrEmpty(ssmsRoot)) throw new InvalidOperationException("EXCELGRID_SSMS_ROOT was not set.");
        var ide = Path.Combine(ssmsRoot, "Common7", "IDE");
        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name = new AssemblyName(args.Name).Name + ".dll";
            foreach (var folder in new[] { ide, Path.Combine(ide, "PublicAssemblies"), Path.Combine(ide, "CommonExtensions", "Microsoft", "Editor"), Path.Combine(ide, "Extensions", "Application") })
            {
                var candidate = Path.Combine(folder, name);
                if (File.Exists(candidate)) return Assembly.LoadFrom(candidate);
            }
            return null;
        };

        var sqlEditors = Assembly.LoadFrom(Path.Combine(ide, "Extensions", "Application", "SQLEditors.dll"));
        var resultSetType = sqlEditors.GetType("Microsoft.SqlServer.Management.QueryExecution.QEResultSet", true);
        var resultSet = Activator.CreateInstance(resultSetType, true);
        var viewField = resultSetType.GetField("m_view", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var storage = new SortableFilterStorage(new FakeStorage(new[] { "works" }));
        var proxy = StorageViewRealProxy.Create(viewField.FieldType, storage);
        viewField.SetValue(resultSet, proxy);
        Equal(true, ReferenceEquals(proxy, viewField.GetValue(resultSet)));

        Equal(1L, (long)resultSetType.GetMethod("NumRows")!.Invoke(resultSet, null));
        Equal("works", (string)resultSetType.GetMethod("GetCellDataAsString")!.Invoke(resultSet, new object[] { 0L, 0 }));

        using var grid = new GridControl();
        grid.AddColumn(new GridColumnInfo());
        grid.GridStorage = (IGridStorage)resultSet;
        grid.UpdateGrid(true);
        var cachedRows = typeof(GridControl).GetProperty("NumRowsInt", BindingFlags.NonPublic | BindingFlags.Instance)!;
        Equal(1L, (long)cachedRows.GetValue(grid, null));

        var nativeRightClickCount = 0;
        var nativeMouseDownCount = 0;
        var nativeFocusedDuringRightClick = false;
        grid.MouseDown += (_, args) =>
        {
            if (args.Button == System.Windows.Forms.MouseButtons.Right) nativeMouseDownCount++;
        };
        grid.MouseButtonClicked += (_, args) =>
        {
            if (args.RowIndex == 0 && args.ColumnIndex == 0 && args.Button == System.Windows.Forms.MouseButtons.Right)
            {
                nativeRightClickCount++;
                nativeFocusedDuringRightClick |= grid.Focused;
            }
        };
        grid.Size = new Size(500, 200);
        grid.CreateControl();
        using (var enhancer = new DevExpressGridEnhancer(grid))
        {
            for (var attempt = 0; attempt < 5; attempt++)
            {
                enhancer.RelayNativeContextMenu(0, 0);
                Equal(false, grid.Capture);
                Equal(true, grid.Focused);
            }
        }
        Equal(5, nativeRightClickCount);
        Equal(5, nativeMouseDownCount);
        Equal(true, nativeFocusedDuringRightClick);
        Equal(false, grid.Capture);
        Equal(true, grid.Focused);

        storage.SetFilter(0, new FilterSpec { AllowedValues = new HashSet<string> { "not present" } });
        resultSetType.GetMethod("OnStorageNotify", BindingFlags.NonPublic | BindingFlags.Instance)!
            .Invoke(resultSet, new object[] { storage.NumRows(), false });
        Equal(0L, (long)resultSetType.GetMethod("NumRows")!.Invoke(resultSet, null));
        Equal(0L, (long)resultSetType.GetField("m_curRowsNum", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(resultSet));
        grid.UpdateGrid(true);
        Equal(0L, (long)cachedRows.GetValue(grid, null));
    }

    private static void DevExpressGridActuallyFiltersRows()
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
        grid.View.ClearColumnsFilter();
        grid.View.ActiveFilterString = "Contains([C0], 'Choc')";
        Equal(1, grid.View.DataRowCount);
        Equal("Heritage Choc", (string)grid.View.GetRowCellValue(0, "C0"));

        grid.View.ClearColumnsFilter();
        grid.View.Columns[0].SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
        var visibleRow = grid.View.GetVisibleRowHandle(0);
        Equal(true, grid.TryMapNativeCell(visibleRow, grid.View.Columns[0], out var nativeRow, out var nativeColumn));
        Equal(3L, nativeRow); // Sorted row still maps to its original SSMS source row.
        Equal(1, nativeColumn);
    }

    private static void DevExpressSurfaceIsHostedInsideNativeGrid()
    {
        using var native = new GridControl { Size = new Size(800, 300), Visible = true, BackColor = Color.White };
        using var enhancer = new DevExpressGridEnhancer(native);
        var replacement = native.Controls.OfType<DevExpressResultsControl>().Single();
        Equal(false, replacement.IsDarkTheme);
        native.BackColor = Color.FromArgb(37, 37, 38);
        enhancer.Refresh();
        Equal(true, replacement.IsDarkTheme);
        Equal(true, ReferenceEquals(native, replacement.Parent));
        Equal(System.Windows.Forms.DockStyle.Fill, replacement.Dock);
        Equal(true, replacement.Visible);
        Equal(true, native.Visible);
        enhancer.Refresh();
        Equal(true, replacement.Visible);
    }

    private static void DialogReturnsUncheckedValues()
    {
        using var dialog = new FilterDialog("name", new[] { "one", "two", "three" }, false, null, true, new Point(100, 100));
        dialog.Show();
        System.Windows.Forms.Application.DoEvents();
        var checks = Descendants(dialog).OfType<System.Windows.Forms.CheckBox>().ToArray();
        Equal(3, checks.Length);
        checks.Single(check => (string)check.Tag == "two").Checked = false;
        var apply = Descendants(dialog).OfType<System.Windows.Forms.Button>().Single(button => button.Text == "Apply");
        apply.PerformClick();
        System.Windows.Forms.Application.DoEvents();
        Equal(true, dialog.Result != null);
        Equal(2, dialog.Result!.AllowedValues!.Count);
        Equal(false, dialog.Result.AllowedValues.Contains("two"));
    }

    private static IEnumerable<System.Windows.Forms.Control> Descendants(System.Windows.Forms.Control parent)
    {
        foreach (System.Windows.Forms.Control child in parent.Controls)
        {
            yield return child;
            foreach (var descendant in Descendants(child)) yield return descendant;
        }
    }
}

internal interface ITestStorageView : IStorageView
{
    int MaxNumBytesToDisplay { get; set; }
}

internal sealed class FakeStorage : ITestStorageView
{
    private readonly string[][] _rows;
    public int MaxNumBytesToDisplay { get; set; }
    public FakeStorage(params string[][] rows) => _rows = rows;
    public long NumRows() => _rows.LongLength;
    public long EnsureRowsInBuf(long first, long last) => last - first + 1;
    public string GetCellDataAsString(long row, int column) => _rows[row][column];
    public object GetCellData(long row, int column) => _rows[row][column];
    public void DeleteRow(long row) => throw new NotSupportedException();
    public IColumnInfo GetColumnInfo(int column) => null!;
    public int NumColumns() => _rows.Length == 0 ? 0 : _rows[0].Length;
    public bool IsStorageClosed() => false;
    public void Dispose() { }
}
