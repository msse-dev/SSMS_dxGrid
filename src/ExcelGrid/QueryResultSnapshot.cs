using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Reflection;
using Microsoft.SqlServer.Management.UI.Grid;

namespace ExcelGrid.Ssms;

internal sealed class QueryResultSnapshot
{
    private const int MaximumRows = 250_000;

    private QueryResultSnapshot(DataTable table, IReadOnlyList<string> captions, long sourceRowCount, bool truncated)
    {
        Table = table;
        Captions = captions;
        SourceRowCount = sourceRowCount;
        Truncated = truncated;
    }

    public DataTable Table { get; }
    public IReadOnlyList<string> Captions { get; }
    public long SourceRowCount { get; }
    public bool Truncated { get; }

    public static QueryResultSnapshot Capture(GridControl grid)
    {
        var storage = grid.GridStorage
            ?? throw new InvalidOperationException("SSMS result storage was not available.");
        var resultSet = (object)storage;
        var getFieldType = resultSet.GetType().GetMethod("GetFieldType", BindingFlags.Public | BindingFlags.Instance);
        var getCellData = resultSet.GetType().GetMethod("GetCellData", BindingFlags.Public | BindingFlags.Instance);

        var table = new DataTable("SSMS Results") { CaseSensitive = false };
        var captions = new List<string>();
        var dataColumns = new List<int>();

        for (var uiColumn = 0; uiColumn < grid.ColumnsNumber; uiColumn++)
        {
            var resultColumn = grid.GetStorageColumnIndexByUIIndex(uiColumn);
            if (resultColumn <= 0) continue; // row-number margin
            var dataColumn = resultColumn - 1;
            string caption = string.Empty;
            Bitmap bitmap;
            try { grid.GetHeaderInfo(uiColumn, out caption, out bitmap); }
            catch { caption = $"Column {dataColumn + 1}"; }
            if (string.IsNullOrWhiteSpace(caption)) caption = $"Column {dataColumn + 1}";

            var fieldType = getFieldType?.Invoke(resultSet, new object[] { dataColumn }) as Type;
            fieldType = NormalizeType(fieldType);
            var column = new DataColumn($"C{dataColumn}", fieldType) { Caption = caption };
            column.ExtendedProperties["SsmsUiColumnIndex"] = uiColumn;
            table.Columns.Add(column);
            captions.Add(caption);
            dataColumns.Add(dataColumn);
        }

        var sourceRows = storage.NumRows();
        var rowsToLoad = Math.Min(sourceRows, MaximumRows);
        const int chunkSize = 2_000;
        for (long start = 0; start < rowsToLoad; start += chunkSize)
        {
            var end = Math.Min(rowsToLoad - 1, start + chunkSize - 1);
            storage.EnsureRowsInBuf(start, end);
            for (var rowIndex = start; rowIndex <= end; rowIndex++)
            {
                var row = table.NewRow();
                for (var index = 0; index < dataColumns.Count; index++)
                {
                    var value = getCellData != null
                        ? getCellData.Invoke(resultSet, new object[] { rowIndex, dataColumns[index] })
                        : storage.GetCellDataAsString(rowIndex, dataColumns[index]);
                    row[index] = NormalizeValue(value, table.Columns[index].DataType);
                }
                table.Rows.Add(row);
            }
        }
        table.AcceptChanges();
        return new QueryResultSnapshot(table, captions, sourceRows, rowsToLoad < sourceRows);
    }

    private static Type NormalizeType(Type? type)
    {
        if (type == null) return typeof(object);
        type = Nullable.GetUnderlyingType(type) ?? type;
        if (type == typeof(byte[])) return typeof(string);
        if (type.IsEnum || type.IsPrimitive || type == typeof(string) || type == typeof(decimal) ||
            type == typeof(DateTime) || type == typeof(DateTimeOffset) || type == typeof(TimeSpan) || type == typeof(Guid))
            return type;
        return typeof(object);
    }

    private static object NormalizeValue(object? value, Type columnType)
    {
        if (value == null || value == DBNull.Value) return DBNull.Value;
        if (value is System.Data.SqlTypes.INullable sqlValue)
        {
            if (sqlValue.IsNull) return DBNull.Value;
            var property = value.GetType().GetProperty("Value");
            if (property != null) value = property.GetValue(value, null);
        }
        if (value is byte[] bytes) return "0x" + BitConverter.ToString(bytes).Replace("-", string.Empty);
        if (columnType == typeof(object) || columnType.IsInstanceOfType(value)) return value;
        try { return Convert.ChangeType(value, columnType); }
        catch { return DBNull.Value; }
    }
}
