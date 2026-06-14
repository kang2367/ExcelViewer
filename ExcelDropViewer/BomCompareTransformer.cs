using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

namespace ExcelDropViewer
{
    internal static class BomCompareTransformer
    {
        private const string MaxTextLengthKey = "MaxTextLength";
        private const string NotFoundValue = "None";

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

        private static readonly string[] CompareSourceHeaders =
        {
            "Part",
            "Voltage",
            "Tolerance",
            "Size"
        };

        private static readonly string[] QuantityHeaderNames =
        {
            "Quantity",
            "Q'ty",
            "Qty",
            "QTY"
        };

        private const string ResultHeaderName = "Result";
        private const string ResultOkValue = "OK";
        private const string ResultNgValue = "NG";
        private const string IsResultColumnKey = "IsResultColumn";

        private static readonly CopyColumnDefinition[] CopyColumnDefinitions =
        {
            new("No", "No", "NO", "No."),
            new("품번", "품번"),
            new("사양", "사양"),
            new("제조사", "제조사"),
            new("Q'ty", "Q'ty", "Quantity", "Qty", "QTY")
        };

        public static DataTable CompareBom(
            DataTable primary,
            DataTable secondary,
            Action<int, int>? onProgress = null)
        {
            if (primary == null || primary.Rows.Count < 2)
            {
                throw new InvalidOperationException("첫 번째 데이터에 비교할 행이 없습니다.");
            }

            if (secondary == null || secondary.Rows.Count < 2)
            {
                throw new InvalidOperationException("두 번째 데이터에 비교할 행이 없습니다.");
            }

            var primaryHeader = primary.Rows[0];
            var secondaryHeader = secondary.Rows[0];

            var compareColumnIndexes = ResolveCompareSourceColumns(primaryHeader);
            var referenceColumn = FindReferenceColumn(primaryHeader);
            var copyColumnMappings = ResolveCopyColumnMappings(secondaryHeader);

            if (referenceColumn < 0)
            {
                throw new InvalidOperationException("첫 번째 데이터의 첫 행에서 Reference 열을 찾을 수 없습니다.");
            }

            if (compareColumnIndexes.Count == 0)
            {
                throw new InvalidOperationException("첫 번째 데이터의 첫 행에서 Part, Voltage, Tolerance, Size 열을 찾을 수 없습니다.");
            }

            var primaryQuantityColumn = FindColumnIndex(primaryHeader, QuantityHeaderNames);
            if (primaryQuantityColumn < 0)
            {
                throw new InvalidOperationException("첫 번째 데이터의 첫 행에서 Quantity 열을 찾을 수 없습니다.");
            }

            var result = primary.Copy();
            var insertIndex = referenceColumn + 1;
            var insertedColumns = new List<DataColumn>();

            for (var offset = 0; offset < copyColumnMappings.Count; offset++)
            {
                var column = result.Columns.Add($"F{result.Columns.Count}", typeof(string));
                column.SetOrdinal(insertIndex + offset);
                column.ExtendedProperties[MaxTextLengthKey] = copyColumnMappings[offset].HeaderName.Length;
                insertedColumns.Add(column);
                result.Rows[0][insertIndex + offset] = copyColumnMappings[offset].HeaderName;
            }

            var specificationColumn = copyColumnMappings.First(m => m.HeaderName == "사양").SourceColumnIndex;
            var partNumberSourceColumn = copyColumnMappings.First(m => m.HeaderName == "품번").SourceColumnIndex;
            var secondaryQuantityColumn = copyColumnMappings.First(m => m.HeaderName == "Q'ty").SourceColumnIndex;

            var resultColumn = result.Columns.Add($"F{result.Columns.Count}", typeof(string));
            resultColumn.ExtendedProperties[MaxTextLengthKey] = ResultHeaderName.Length;
            resultColumn.ExtendedProperties[IsResultColumnKey] = true;
            var resultColumnIndex = resultColumn.Ordinal;
            result.Rows[0][resultColumnIndex] = ResultHeaderName;

            var totalRows = Math.Max(0, result.Rows.Count - 1);
            var processedRows = 0;

            for (var rowIndex = 1; rowIndex < result.Rows.Count; rowIndex++)
            {
                processedRows++;
                onProgress?.Invoke(processedRows, totalRows);

                var tokens = CollectCompareTokens(result.Rows[rowIndex], compareColumnIndexes);
                var matchedRowIndex = tokens.Count == 0
                    ? -1
                    : FindMatchingRowIndex(secondary, specificationColumn, partNumberSourceColumn, tokens);

                for (var columnOffset = 0; columnOffset < copyColumnMappings.Count; columnOffset++)
                {
                    var value = matchedRowIndex < 0
                        ? NotFoundValue
                        : GetCopiedCellValue(secondary.Rows[matchedRowIndex], copyColumnMappings[columnOffset].SourceColumnIndex);

                    result.Rows[rowIndex][insertIndex + columnOffset] = value;
                    UpdateMaxTextLength(insertedColumns[columnOffset], value);
                }

                if (matchedRowIndex < 0)
                {
                    result.Rows[rowIndex][resultColumnIndex] = string.Empty;
                    continue;
                }

                var primaryQuantity = GetCellText(result.Rows[rowIndex][primaryQuantityColumn]);
                var secondaryQuantity = GetCellText(secondary.Rows[matchedRowIndex][secondaryQuantityColumn]);
                var compareResult = QuantitiesMatch(primaryQuantity, secondaryQuantity)
                    ? ResultOkValue
                    : ResultNgValue;

                result.Rows[rowIndex][resultColumnIndex] = compareResult;
                UpdateMaxTextLength(resultColumn, compareResult);
            }

            return result;
        }

        private static bool QuantitiesMatch(string primaryQuantity, string secondaryQuantity)
        {
            if (TryParseQuantity(primaryQuantity, out var primaryValue)
                && TryParseQuantity(secondaryQuantity, out var secondaryValue))
            {
                return primaryValue == secondaryValue;
            }

            return string.Equals(primaryQuantity, secondaryQuantity, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryParseQuantity(string text, out decimal value)
        {
            if (decimal.TryParse(text, NumberStyles.Number, CultureInfo.CurrentCulture, out value))
            {
                return true;
            }

            return decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out value);
        }

        private sealed record CopyColumnDefinition(string HeaderName, params string[] HeaderAliases);

        private sealed class CopyColumnMapping
        {
            public CopyColumnMapping(string headerName, int sourceColumnIndex)
            {
                HeaderName = headerName;
                SourceColumnIndex = sourceColumnIndex;
            }

            public string HeaderName { get; }

            public int SourceColumnIndex { get; }
        }

        private static List<CopyColumnMapping> ResolveCopyColumnMappings(DataRow secondaryHeader)
        {
            var mappings = new List<CopyColumnMapping>();

            foreach (var definition in CopyColumnDefinitions)
            {
                var columnIndex = FindColumnIndex(secondaryHeader, definition.HeaderAliases);
                if (columnIndex < 0)
                {
                    throw new InvalidOperationException(
                        $"두 번째 데이터의 첫 행에서 {definition.HeaderName} 열을 찾을 수 없습니다.");
                }

                mappings.Add(new CopyColumnMapping(definition.HeaderName, columnIndex));
            }

            return mappings;
        }

        private static void UpdateMaxTextLength(DataColumn column, string value)
        {
            var currentMax = column.ExtendedProperties[MaxTextLengthKey] is int max ? max : 0;
            if (value.Length > currentMax)
            {
                column.ExtendedProperties[MaxTextLengthKey] = value.Length;
            }
        }

        private static string GetCopiedCellValue(DataRow sourceRow, int sourceColumnIndex)
        {
            var value = GetCellText(sourceRow[sourceColumnIndex]);
            return string.IsNullOrWhiteSpace(value) ? NotFoundValue : value;
        }

        private static List<int> ResolveCompareSourceColumns(DataRow headerRow)
        {
            var indexes = new List<int>();
            foreach (var headerName in CompareSourceHeaders)
            {
                var columnIndex = FindColumnIndex(headerRow, headerName);
                if (columnIndex >= 0)
                {
                    indexes.Add(columnIndex);
                }
            }

            return indexes;
        }

        private static int FindReferenceColumn(DataRow headerRow)
        {
            for (var columnIndex = 0; columnIndex < headerRow.Table.Columns.Count; columnIndex++)
            {
                if (MatchesReferenceHeader(GetCellText(headerRow[columnIndex])))
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

        private static int FindColumnIndex(DataRow headerRow, params string[] headerNames)
        {
            foreach (var headerName in headerNames)
            {
                for (var columnIndex = 0; columnIndex < headerRow.Table.Columns.Count; columnIndex++)
                {
                    if (GetCellText(headerRow[columnIndex]).Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    {
                        return columnIndex;
                    }
                }
            }

            return -1;
        }

        private static List<string> CollectCompareTokens(DataRow row, IReadOnlyList<int> columnIndexes)
        {
            var tokens = new List<string>();
            foreach (var columnIndex in columnIndexes)
            {
                var value = GetCellText(row[columnIndex]);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    tokens.Add(value);
                }
            }

            return tokens;
        }

        private static int FindMatchingRowIndex(
            DataTable secondary,
            int specificationColumn,
            int partNumberColumn,
            IReadOnlyList<string> tokens)
        {
            for (var rowIndex = 1; rowIndex < secondary.Rows.Count; rowIndex++)
            {
                var row = secondary.Rows[rowIndex];
                var specification = GetCellText(row[specificationColumn]);
                var partNumberText = GetCellText(row[partNumberColumn]);

                if (MatchesTokensInSearchColumns(specification, partNumberText, tokens))
                {
                    return rowIndex;
                }
            }

            return -1;
        }

        private static bool MatchesTokensInSearchColumns(
            string specification,
            string partNumber,
            IReadOnlyList<string> tokens)
        {
            foreach (var token in tokens)
            {
                var foundInSpecification = !string.IsNullOrWhiteSpace(specification)
                    && specification.Contains(token, StringComparison.OrdinalIgnoreCase);
                var foundInPartNumber = !string.IsNullOrWhiteSpace(partNumber)
                    && partNumber.Contains(token, StringComparison.OrdinalIgnoreCase);

                if (!foundInSpecification && !foundInPartNumber)
                {
                    return false;
                }
            }

            return true;
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
