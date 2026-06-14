using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

namespace ExcelDropViewer
{
    internal static class BomOneRowTransformer
    {
        private const string MaxTextLengthKey = "MaxTextLength";

        private static readonly string[] ItemHeaderNames =
        {
            "Item",
            "아이템"
        };

        private static readonly string[] ReferenceHeaderNames =
        {
            "Reference",
            "Reference Designator",
            "Ref. Designator",
            "Ref Designator",
            "Ref",
            "Ref Des",
            "Designator",
            "부품위치"
        };

        public static DataTable TransformWithSelectedHeaderRow(
            DataTable source,
            int headerRowIndex,
            Action<int, int>? onProgress = null)
        {
            if (source == null || source.Rows.Count == 0)
            {
                throw new InvalidOperationException("변환할 데이터가 없습니다.");
            }

            if (headerRowIndex < 0 || headerRowIndex >= source.Rows.Count)
            {
                throw new InvalidOperationException("선택한 헤더 행 위치가 유효하지 않습니다.");
            }

            if (headerRowIndex + 3 >= source.Rows.Count)
            {
                throw new InvalidOperationException("헤더 행 기준 변환할 데이터 행이 없습니다.");
            }

            var virtualHeaders = ExtractVirtualHeaders(source.Rows[headerRowIndex], source.Columns.Count);
            var itemColumn = FindItemColumn(virtualHeaders);
            var referenceColumn = FindReferenceColumn(virtualHeaders);

            if (itemColumn < 0 || referenceColumn < 0)
            {
                throw new InvalidOperationException(
                    "선택한 행에서 Item 또는 Reference 열을 찾을 수 없습니다.");
            }

            var result = CreateEmptyTableLike(source);

            var firstRow = result.NewRow();
            PopulateFirstRow(firstRow, source.Rows[headerRowIndex], virtualHeaders, itemColumn);
            result.Rows.Add(firstRow);

            DataRow? mergedRow = null;
            var dataStartRow = headerRowIndex + 3;
            var totalRows = Math.Max(0, source.Rows.Count - dataStartRow);
            var processedRows = 0;

            for (var rowIndex = dataStartRow; rowIndex < source.Rows.Count; rowIndex++)
            {
                processedRows++;
                onProgress?.Invoke(processedRows, totalRows);

                var sourceRow = source.Rows[rowIndex];
                var itemValue = GetCellText(sourceRow[itemColumn]);

                if (!string.IsNullOrWhiteSpace(itemValue))
                {
                    if (mergedRow != null)
                    {
                        result.Rows.Add(mergedRow);
                    }

                    mergedRow = result.NewRow();
                    CopyRowValues(sourceRow, mergedRow);
                    continue;
                }

                if (mergedRow == null)
                {
                    var orphanRow = result.NewRow();
                    CopyRowValues(sourceRow, orphanRow);
                    result.Rows.Add(orphanRow);
                    continue;
                }

                AppendReference(mergedRow, referenceColumn, GetCellText(sourceRow[referenceColumn]));
                AppendOtherColumnValues(mergedRow, sourceRow, itemColumn, referenceColumn);
            }

            if (mergedRow != null)
            {
                result.Rows.Add(mergedRow);
            }

            RecalculateMaxTextLengths(result);
            return result;
        }

        private static string[] ExtractVirtualHeaders(DataRow headerRow, int columnCount)
        {
            var headers = new string[columnCount];
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                headers[columnIndex] = GetCellText(headerRow[columnIndex]);
            }

            return headers;
        }

        private static int FindItemColumn(IReadOnlyList<string> columnHeaders)
        {
            for (var columnIndex = 0; columnIndex < columnHeaders.Count; columnIndex++)
            {
                if (MatchesHeader(columnHeaders[columnIndex], ItemHeaderNames))
                {
                    return columnIndex;
                }
            }

            return -1;
        }

        private static int FindReferenceColumn(IReadOnlyList<string> columnHeaders)
        {
            for (var columnIndex = 0; columnIndex < columnHeaders.Count; columnIndex++)
            {
                if (MatchesReferenceHeader(columnHeaders[columnIndex]))
                {
                    return columnIndex;
                }
            }

            return -1;
        }

        private static bool MatchesReferenceHeader(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var normalized = text.Trim();
            foreach (var candidate in ReferenceHeaderNames.OrderByDescending(static name => name.Length))
            {
                if (normalized.Equals(candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool MatchesHeader(string text, IReadOnlyList<string> candidates)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var normalized = text.Trim();
            foreach (var candidate in candidates)
            {
                if (normalized.Equals(candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static DataTable CreateEmptyTableLike(DataTable source)
        {
            var table = new DataTable();
            foreach (DataColumn sourceColumn in source.Columns)
            {
                var column = table.Columns.Add(sourceColumn.ColumnName, typeof(string));
                column.ExtendedProperties[MaxTextLengthKey] = 0;
            }

            return table;
        }

        private static void PopulateFirstRow(
            DataRow firstRow,
            DataRow selectedRow,
            IReadOnlyList<string> virtualHeaders,
            int itemColumn)
        {
            for (var columnIndex = 0; columnIndex < virtualHeaders.Count; columnIndex++)
            {
                var selectedValue = GetCellText(selectedRow[columnIndex]);
                firstRow[columnIndex] = !string.IsNullOrWhiteSpace(selectedValue)
                    ? selectedValue
                    : virtualHeaders[columnIndex];
            }

            if (string.IsNullOrWhiteSpace(GetCellText(firstRow[itemColumn])))
            {
                firstRow[itemColumn] = virtualHeaders[itemColumn];
            }

            if (itemColumn == 0 && string.IsNullOrWhiteSpace(GetCellText(firstRow[0])))
            {
                firstRow[0] = "Item";
            }
            else if (string.IsNullOrWhiteSpace(GetCellText(firstRow[0]))
                     && !string.IsNullOrWhiteSpace(virtualHeaders[itemColumn]))
            {
                firstRow[0] = virtualHeaders[itemColumn];
            }
        }

        private static void CopyRowValues(DataRow source, DataRow target)
        {
            for (var columnIndex = 0; columnIndex < source.Table.Columns.Count; columnIndex++)
            {
                target[columnIndex] = GetCellText(source[columnIndex]);
            }
        }

        private static void AppendReference(DataRow mergedRow, int referenceColumn, string referenceValue)
        {
            if (string.IsNullOrWhiteSpace(referenceValue))
            {
                return;
            }

            var existing = GetCellText(mergedRow[referenceColumn]);
            mergedRow[referenceColumn] = string.IsNullOrWhiteSpace(existing)
                ? referenceValue.Trim()
                : $"{existing.Trim()}\n{referenceValue.Trim()}";
        }

        private static void AppendOtherColumnValues(
            DataRow mergedRow,
            DataRow continuationRow,
            int itemColumn,
            int referenceColumn)
        {
            for (var columnIndex = 0; columnIndex < continuationRow.Table.Columns.Count; columnIndex++)
            {
                if (columnIndex == itemColumn || columnIndex == referenceColumn)
                {
                    continue;
                }

                var value = GetCellText(continuationRow[columnIndex]);
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                var existing = GetCellText(mergedRow[columnIndex]);
                if (string.IsNullOrWhiteSpace(existing))
                {
                    mergedRow[columnIndex] = value;
                    continue;
                }

                if (!existing.Contains(value, StringComparison.Ordinal))
                {
                    mergedRow[columnIndex] = $"{existing}\n{value}";
                }
            }
        }

        private static void RecalculateMaxTextLengths(DataTable table)
        {
            foreach (DataColumn column in table.Columns)
            {
                column.ExtendedProperties[MaxTextLengthKey] = 0;
            }

            foreach (DataRow row in table.Rows)
            {
                for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    var text = GetCellText(row[columnIndex]);
                    var column = table.Columns[columnIndex];
                    var currentMax = column.ExtendedProperties[MaxTextLengthKey] is int max ? max : 0;
                    if (text.Length > currentMax)
                    {
                        column.ExtendedProperties[MaxTextLengthKey] = text.Length;
                    }
                }
            }
        }

        private static string GetCellText(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            return Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
        }
    }
}
