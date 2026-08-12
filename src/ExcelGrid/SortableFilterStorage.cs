using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Microsoft.SqlServer.Management.UI.Grid;

namespace ExcelGrid.Ssms;

internal enum SortDirection { None, Ascending, Descending }

/// <summary>
/// Remaps an SSMS storage view while leaving QEResultSet itself installed on the GridControl.
/// SSMS casts GridStorage back to QEResultSet in selection/copy paths, so replacing GridStorage
/// with an IGridStorage wrapper is unsafe.
/// </summary>
internal sealed class SortableFilterStorage
{
    private readonly IStorageView _inner;
    private readonly object _innerObject;
    private readonly PropertyInfo? _maxBytesProperty;
    private readonly Dictionary<int, FilterSpec> _filters = new();
    private long[] _rows = Array.Empty<long>();
    private long _sourceRowCount = -1;

    public SortableFilterStorage(object inner)
    {
        _innerObject = inner;
        _inner = (IStorageView)inner;
        _maxBytesProperty = inner.GetType().GetProperty("MaxNumBytesToDisplay", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
        Rebuild();
    }

    public int SortColumn { get; private set; } = -1;
    public SortDirection SortDirection { get; private set; }
    public IReadOnlyDictionary<int, FilterSpec> Filters => _filters;
    public long SourceRowCount => _sourceRowCount;

    // Required by SSMS's internal IQEStorageView interface. The transparent proxy forwards it here.
    public int MaxNumBytesToDisplay
    {
        get => _maxBytesProperty == null ? 0 : (int)_maxBytesProperty.GetValue(_innerObject, null);
        set => _maxBytesProperty?.SetValue(_innerObject, value, null);
    }

    public SortDirection CycleSort(int column)
    {
        if (SortColumn != column)
        {
            SortColumn = column;
            SortDirection = SortDirection.Ascending;
        }
        else
        {
            SortDirection = SortDirection switch
            {
                SortDirection.Ascending => SortDirection.Descending,
                SortDirection.Descending => SortDirection.None,
                _ => SortDirection.Ascending
            };
        }
        Rebuild();
        return SortDirection;
    }

    public void SetFilter(int column, FilterSpec? filter)
    {
        if (filter == null || !filter.IsActive)
            _filters.Remove(column);
        else
            _filters[column] = filter;
        Rebuild();
    }

    public IReadOnlyList<string> GetDistinctValues(int column, int limit, out bool truncated)
    {
        var values = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
        var count = _inner.NumRows();
        truncated = false;
        for (long row = 0; row < count; row++)
        {
            values.Add(_inner.GetCellDataAsString(row, column) ?? string.Empty);
            if (values.Count > limit)
            {
                truncated = true;
                break;
            }
        }
        return values.Take(limit).OrderBy(x => x, NaturalStringComparer.Instance).ToArray();
    }

    private void Rebuild()
    {
        var rows = new List<long>();
        var count = _inner.NumRows();
        for (long row = 0; row < count; row++)
        {
            var keep = true;
            foreach (var pair in _filters)
            {
                if (!pair.Value.Matches(_inner.GetCellDataAsString(row, pair.Key)))
                {
                    keep = false;
                    break;
                }
            }
            if (keep) rows.Add(row);
        }

        if (SortDirection != SortDirection.None && SortColumn >= 0)
        {
            rows.Sort((left, right) =>
            {
                var result = NaturalStringComparer.Instance.Compare(
                    _inner.GetCellDataAsString(left, SortColumn),
                    _inner.GetCellDataAsString(right, SortColumn));
                if (result != 0 && SortDirection == SortDirection.Descending) result = -result;
                return result == 0 ? left.CompareTo(right) : result;
            });
        }
        _rows = rows.ToArray();
        _sourceRowCount = count;
    }

    private void RefreshIfRowsArrived()
    {
        if (_inner.NumRows() != _sourceRowCount) Rebuild();
    }

    private long Map(long row)
    {
        RefreshIfRowsArrived();
        return row >= 0 && row < _rows.LongLength ? _rows[row] : row;
    }

    public long NumRows()
    {
        RefreshIfRowsArrived();
        return _rows.LongLength;
    }

    public long EnsureRowsInBuf(long firstRow, long lastRow)
    {
        RefreshIfRowsArrived();
        if (_rows.Length == 0) return 0;
        var first = Math.Max(0, firstRow);
        var last = Math.Min(_rows.LongLength - 1, lastRow);
        if (last < first) return 0;
        long min = long.MaxValue, max = long.MinValue;
        for (var index = first; index <= last; index++)
        {
            var mapped = Map(index);
            min = Math.Min(min, mapped);
            max = Math.Max(max, mapped);
        }
        return _inner.EnsureRowsInBuf(min, max);
    }

    public object GetCellData(long row, int column) => _inner.GetCellData(Map(row), column);
    public string GetCellDataAsString(long row, int column) => _inner.GetCellDataAsString(Map(row), column);
    public IColumnInfo GetColumnInfo(int column) => _inner.GetColumnInfo(column);
    public int NumColumns() => _inner.NumColumns();
    public bool IsStorageClosed() => _inner.IsStorageClosed();

    public void DeleteRow(long row)
    {
        _inner.DeleteRow(Map(row));
        Rebuild();
    }

    public void Dispose() => _inner.Dispose();
}

/// <summary>
/// IQEStorageView is internal to SQLEditors.dll. RealProxy lets us implement that exact
/// runtime interface without taking a compile-time dependency on a private SSMS type.
/// </summary>
internal sealed class StorageViewRealProxy : RealProxy
{
    private readonly SortableFilterStorage _target;

    private StorageViewRealProxy(Type interfaceType, SortableFilterStorage target) : base(interfaceType) => _target = target;

    public static object Create(Type interfaceType, SortableFilterStorage target) =>
        new StorageViewRealProxy(interfaceType, target).GetTransparentProxy();

    public override IMessage Invoke(IMessage message)
    {
        var call = (IMethodCallMessage)message;
        try
        {
            var parameterTypes = call.MethodBase.GetParameters().Select(p => p.ParameterType).ToArray();
            var method = typeof(SortableFilterStorage).GetMethod(
                call.MethodName,
                BindingFlags.Public | BindingFlags.Instance,
                null,
                parameterTypes,
                null) ?? throw new MissingMethodException(typeof(SortableFilterStorage).FullName, call.MethodName);
            var result = method.Invoke(_target, call.Args);
            return new ReturnMessage(result, null, 0, call.LogicalCallContext, call);
        }
        catch (TargetInvocationException error)
        {
            return new ReturnMessage(error.InnerException ?? error, call);
        }
        catch (Exception error)
        {
            return new ReturnMessage(error, call);
        }
    }
}

internal sealed class NaturalStringComparer : IComparer<string>
{
    public static readonly NaturalStringComparer Instance = new();

    public int Compare(string? x, string? y)
    {
        x ??= string.Empty;
        y ??= string.Empty;
        if (decimal.TryParse(x, NumberStyles.Any, CultureInfo.CurrentCulture, out var xd) &&
            decimal.TryParse(y, NumberStyles.Any, CultureInfo.CurrentCulture, out var yd))
            return xd.CompareTo(yd);
        if (DateTime.TryParse(x, CultureInfo.CurrentCulture, DateTimeStyles.None, out var xt) &&
            DateTime.TryParse(y, CultureInfo.CurrentCulture, DateTimeStyles.None, out var yt))
            return xt.CompareTo(yt);
        return StringComparer.CurrentCultureIgnoreCase.Compare(x, y);
    }
}
