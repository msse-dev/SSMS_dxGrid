using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.SqlServer.Management.UI.Grid;

namespace ExcelGrid.Ssms;

internal static class GridDiscovery
{
    private static readonly Dictionary<GridControl, DevExpressGridEnhancer> Enhancers = new();
    private static readonly HashSet<GridControl> CandidateGrids = new();
    private static readonly HashSet<Control> WatchedControls = new();
    private static Timer? _timer;

    public static void Start()
    {
        if (_timer != null) return;
        // SSMS can create a fresh results grid for each execution. Poll quickly enough
        // to attach before the native surface can visibly flash or require interaction.
        _timer = new Timer { Interval = 100 };
        _timer.Tick += (_, _) => Scan();
        _timer.Start();
        Scan();
    }

    private static void Scan()
    {
        // A future SSMS build may create a query grid before assigning QEResultSet.
        // Recheck unattached candidates so the storage marker can qualify it later.
        foreach (var grid in new List<GridControl>(CandidateGrids))
        {
            if (grid.IsDisposed)
                CandidateGrids.Remove(grid);
            else
                Attach(grid);
        }

        foreach (var pair in Enhancers)
        {
            if (!pair.Value.IsDisposed)
                pair.Value.Refresh();
        }

        var processId = (uint)Process.GetCurrentProcess().Id;
        EnumWindows((topLevel, _) =>
        {
            GetWindowThreadProcessId(topLevel, out var ownerProcessId);
            if (ownerProcessId != processId) return true;

            InspectHandle(topLevel);
            EnumChildWindows(topLevel, (handle, __) =>
            {
                InspectHandle(handle);
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);

        var stale = new List<GridControl>();
        foreach (var pair in Enhancers)
            if (pair.Value.IsDisposed) stale.Add(pair.Key);
        foreach (var grid in stale)
        {
            Enhancers[grid].Dispose();
            Enhancers.Remove(grid);
        }
    }

    private static void InspectHandle(IntPtr handle)
    {
        var control = Control.FromChildHandle(handle);
        if (control != null) WatchTree(control);
    }

    private static void WatchTree(Control control)
    {
        // Do not recursively observe our own DevExpress tree.
        if (control is DevExpressResultsControl) return;
        if (!WatchedControls.Add(control)) return;

        control.ControlAdded += ControlAdded;
        control.Disposed += WatchedControlDisposed;
        if (control is GridControl grid)
        {
            CandidateGrids.Add(grid);
            Attach(grid);
        }
        foreach (Control child in control.Controls) WatchTree(child);
    }

    private static void ControlAdded(object sender, ControlEventArgs e)
    {
        // ControlAdded is synchronous with SSMS constructing its result surface, so the
        // replacement child is installed before the native grid's first paint.
        WatchTree(e.Control);
    }

    private static void WatchedControlDisposed(object sender, EventArgs e)
    {
        if (sender is Control control)
        {
            WatchedControls.Remove(control);
            if (control is GridControl grid) CandidateGrids.Remove(grid);
        }
    }

    private static void Attach(GridControl grid)
    {
        if (grid.IsDisposed || !IsQueryResultsGrid(grid)) return;
        if (!Enhancers.TryGetValue(grid, out var enhancer))
            Enhancers[grid] = new DevExpressGridEnhancer(grid);
        else
            enhancer.Refresh();
    }

    internal static bool IsQueryResultsGrid(GridControl grid)
    {
        const string gridTypeName =
            "Microsoft.SqlServer.Management.UI.VSIntegration.Editors.GridResultsGrid";
        const string gridInterfaceName =
            "Microsoft.SqlServer.Management.QueryExecution.IGridControl2";
        const string storageTypeName =
            "Microsoft.SqlServer.Management.QueryExecution.QEResultSet";
        const string storageInterfaceName =
            "Microsoft.SqlServer.Management.QueryExecution.IGridResultSet";

        var gridType = grid.GetType();
        if (gridType.FullName == gridTypeName || Implements(gridType, gridInterfaceName)) return true;

        var storageType = grid.GridStorage?.GetType();
        return storageType != null &&
               (storageType.FullName == storageTypeName || Implements(storageType, storageInterfaceName));
    }

    private static bool Implements(Type type, string interfaceName)
    {
        foreach (var contract in type.GetInterfaces())
            if (contract.FullName == interfaceName) return true;
        return false;
    }

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}

internal sealed class DevExpressGridEnhancer : IDisposable
{
    private readonly GridControl _native;
    private readonly DevExpressResultsControl? _replacement;
    private object? _lastStorage;
    private long _lastRows = -1;
    private bool _building;

    private const int WmRightButtonDown = 0x0204;
    private const int MkRightButton = 0x0002;

    public DevExpressGridEnhancer(GridControl native)
    {
        _native = native;
        _native.Disposed += NativeDisposed;
        _replacement = new DevExpressResultsControl(SsmsTheme.IsDark(native))
        {
            Visible = true,
            Dock = DockStyle.Fill
        };
        // Host inside the SSMS grid rather than beside it. SSMS frequently calls
        // BringToFront on its native result control while executing/focusing tabs;
        // a child surface cannot be placed behind its own parent.
        native.Controls.Add(_replacement);
        _replacement.NativeContextMenuRequested += ShowNativeContextMenu;
        _replacement.BringToFront();
        _replacement.ShowLoading("Waiting for query results…");
        Refresh();
    }

    public bool IsDisposed => _native.IsDisposed;

    public void Refresh()
    {
        if (_native.IsDisposed || _replacement == null || _replacement.IsDisposed || _building) return;
        _replacement.ApplyTheme(SsmsTheme.IsDark(_native));
        SyncLayout();

        var storage = _native.GridStorage;
        if (storage == null || _native.ColumnsNumber <= 1)
        {
            _replacement.ShowLoading("Running query…");
            ShowReplacement();
            return;
        }

        var resultSet = (object)storage;
        var storedAll = resultSet.GetType().GetProperty("StoredAllData", BindingFlags.Public | BindingFlags.Instance);
        if (storedAll != null && !(bool)storedAll.GetValue(resultSet, null))
        {
            _replacement.ShowLoading("Running query…");
            ShowReplacement();
            return;
        }

        var rows = storage.NumRows();
        if (ReferenceEquals(storage, _lastStorage) && rows == _lastRows)
        {
            if (!_replacement.Visible) ShowReplacement();
            return;
        }

        _building = true;
        try
        {
            var snapshot = QueryResultSnapshot.Capture(_native);
            _replacement.SetData(snapshot.Table, snapshot.Captions, snapshot.SourceRowCount, snapshot.Truncated);
            _lastStorage = storage;
            _lastRows = rows;
            ShowReplacement();
        }
        catch (Exception error)
        {
            _replacement.ShowError(GetUsefulErrorMessage(error));
            ShowReplacement();
        }
        finally
        {
            _building = false;
        }
    }

    private void SyncLayout()
    {
        if (_replacement == null) return;
        if (!ReferenceEquals(_replacement.Parent, _native))
        {
            _native.Controls.Add(_replacement);
            _replacement.BringToFront();
        }
        _replacement.Dock = DockStyle.Fill;
        _replacement.Bounds = _native.ClientRectangle;
    }

    private void ShowReplacement()
    {
        if (_replacement == null) return;
        // Keep SSMS's grid alive as the host and keep our child at the top of its child
        // z-order. Native painting remains underneath the child HWND.
        _replacement.Visible = true;
        _replacement.BringToFront();
    }

    private void ShowNativeContextMenu(object sender, NativeContextMenuRequestEventArgs e)
        => RelayNativeContextMenu(e.Row, e.Column, e.SelectedCells);

    internal void RelayNativeContextMenu(long row, int column, IReadOnlyList<NativeCell>? selectedCells = null)
    {
        if (_native.IsDisposed || !_native.IsHandleCreated) return;
        try
        {
            // Keep SSMS's selection model in sync. SQL Prompt reads this native state
            // when it contributes commands to the standard results-grid menu.
            var selected = BuildNativeSelection(selectedCells);
            if (selected.Count == 0) selected.Add(new BlockOfCells(row, column));
            var onlyBlock = selected.Count == 1 ? selected[0] : null;
            _native.SelectionType = onlyBlock != null && onlyBlock.Width == 1 && onlyBlock.Height == 1
                ? GridSelectionType.SingleCell
                : GridSelectionType.CellBlocks;
            _native.SelectedCells = selected;
            _native.EnsureCellIsVisible(row, column);

            var rectangle = _native.GetVisibleCellRectangle(row, column);
            var x = rectangle.Left + Math.Max(1, rectangle.Width / 2);
            var y = rectangle.Top + Math.Max(1, rectangle.Height / 2);

            // SQL Prompt participates through Visual Studio command routing and only
            // contributes its items when the native SSMS grid is the focused target.
            _native.Select();
            _native.Focus();
            SetFocus(_native.Handle);

            // SSMS dispatches its extensible results-menu event during WM_RBUTTONDOWN.
            // Do not synthesize WM_RBUTTONUP: SendMessage does not return until the
            // popup's modal loop ends, so an "up" sent afterward is stale and makes
            // the native grid's next context-menu invocation unreliable.
            SendMessage(_native.Handle, WmRightButtonDown, new IntPtr(MkRightButton), MakeLParam(x, y));
        }
        catch
        {
            // A results tab can disappear between the DevExpress click and this relay.
        }
        finally
        {
            if (!_native.IsDisposed) _native.Capture = false;
            ResetNativeCaptureTracker();
            ShowReplacement();
        }
    }

    internal static BlockOfCellsCollection BuildNativeSelection(IReadOnlyList<NativeCell>? cells)
    {
        var result = new BlockOfCellsCollection();
        if (cells == null || cells.Count == 0) return result;

        // DevExpress reports one item per selected cell. Passing those directly to
        // SSMS makes a select-all operation thousands of one-cell blocks: its menu
        // becomes slow and Copy flattens the blocks. Preserve display-row order,
        // collapse each row into column runs, then merge consecutive native rows.
        var rows = new List<KeyValuePair<long, SortedSet<int>>>();
        var rowIndexes = new Dictionary<long, int>();
        foreach (var cell in cells)
        {
            if (cell.Row < 0 || cell.Column < 0) continue;
            if (!rowIndexes.TryGetValue(cell.Row, out var rowIndex))
            {
                rowIndex = rows.Count;
                rowIndexes.Add(cell.Row, rowIndex);
                rows.Add(new KeyValuePair<long, SortedSet<int>>(cell.Row, new SortedSet<int>()));
            }
            rows[rowIndex].Value.Add(cell.Column);
        }

        var previousRuns = new Dictionary<(int Left, int Right), BlockOfCells>();
        foreach (var row in rows)
        {
            var runs = GetColumnRuns(row.Value);
            var currentRuns = new Dictionary<(int Left, int Right), BlockOfCells>();
            foreach (var run in runs)
            {
                if (previousRuns.TryGetValue(run, out var existing) && existing.Bottom + 1 == row.Key)
                {
                    existing.Height++;
                    currentRuns.Add(run, existing);
                    continue;
                }

                var block = new BlockOfCells(row.Key, run.Left)
                {
                    Width = run.Right - run.Left + 1
                };
                result.Add(block);
                currentRuns.Add(run, block);
            }
            previousRuns = currentRuns;
        }
        return result;
    }

    private static List<(int Left, int Right)> GetColumnRuns(SortedSet<int> columns)
    {
        var runs = new List<(int Left, int Right)>();
        var hasRun = false;
        var left = 0;
        var right = 0;
        foreach (var column in columns)
        {
            if (!hasRun)
            {
                left = right = column;
                hasRun = true;
            }
            else if (column == right + 1)
            {
                right = column;
            }
            else
            {
                runs.Add((left, right));
                left = right = column;
            }
        }
        if (hasRun) runs.Add((left, right));
        return runs;
    }

    private void ResetNativeCaptureTracker()
    {
        try
        {
            var tracker = typeof(GridControl).GetField("m_captureTracker", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.GetValue(_native);
            tracker?.GetType().GetMethod("Reset", BindingFlags.Public | BindingFlags.Instance)?.Invoke(tracker, null);
        }
        catch { }
    }

    private static string GetUsefulErrorMessage(Exception error)
    {
        while (error is TargetInvocationException && error.InnerException != null)
            error = error.InnerException;
        return error.Message;
    }

    private static IntPtr MakeLParam(int x, int y) =>
        new IntPtr((y & 0xffff) << 16 | (x & 0xffff));

    private void ShowNative()
    {
        if (_replacement != null) _replacement.Visible = false;
        if (!_native.IsDisposed) _native.Visible = true;
    }

    private void NativeDisposed(object sender, EventArgs e) => Dispose();

    public void Dispose()
    {
        if (!_native.IsDisposed)
        {
            _native.Disposed -= NativeDisposed;
            _native.Visible = true;
        }
        if (_replacement != null) _replacement.NativeContextMenuRequested -= ShowNativeContextMenu;
        _replacement?.Dispose();
    }

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr window);

}
