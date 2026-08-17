using System.Data;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelDropViewer
{
    internal sealed class PdbBomDbMatchResult
    {
        public DataTable OutputTable { get; init; } = new();
        public PdbBomDbMatchLayout Layout { get; init; } = new();
        public int TotalProcessed { get; init; }
        public int MatchedCount { get; init; }
        public int UnmatchedCount { get; init; }
    }

    internal sealed class PdbBomDbMatchLayout
    {
        public int HeaderRow { get; init; }
        public int DataStartRow { get; init; }
        public int DataEndRow { get; init; }
        public int PartNameColumn { get; init; } = -1;
        public int PartNumberColumn { get; init; } = -1;
        public int ManufacturerColumn { get; init; } = -1;
        public int DbDescriptionColumn { get; init; } = -1;
        public IReadOnlyList<int> UnmatchedRows { get; init; } = Array.Empty<int>();
    }

    internal static class PdbBomDbMatcher
    {
        private const string UnregisteredPartText = "미등록 부품";

        private static readonly Regex SeparatorTextPattern = new(
            @"^[-_=~*.·•─━┅┄┈┉―﹣－\s]+$",
            RegexOptions.Compiled);

        private static readonly string[] ItemDescriptionHeaders = { "품목 설명", "품목설명", "Description" };
        private static readonly string[] PartNameHeaders = { "품명", "Part Name", "Item Name", "품목" };
        private static readonly string[] PartNumberHeaders = { "품목 번호", "품목번호", "Part Number", "Item Code", "Part No" };
        private static readonly string[] ManufacturerHeaders = { "제조사", "Manufacturer", "Mfr" };
        private static readonly string[] DbDescriptionHeaders = { "DB 품목 설명", "DB Description" };

        public static PdbBomDbMatchResult Match(
            DataTable sourceTable,
            string databasePath,
            Action<int, int>? onProgress = null)
        {
            if (sourceTable.Rows.Count == 0)
            {
                throw new InvalidOperationException("매칭할 BOM 데이터가 없습니다.");
            }

            if (!File.Exists(databasePath))
            {
                throw new FileNotFoundException("Data/BOM_Master.db 파일을 찾을 수 없습니다.", databasePath);
            }

            var headerRowIndex = FindItemDescriptionHeaderRow(sourceTable);
            if (headerRowIndex < 0)
            {
                throw new InvalidOperationException(
                    "BOM 파일에서 '품목 설명' 또는 'Description' 헤더 행을 찾을 수 없습니다.");
            }

            var outputTable = sourceTable.Copy();
            var headerRow = outputTable.Rows[headerRowIndex];
            var itemDescriptionColumn = FindItemDescriptionColumn(headerRow);
            if (itemDescriptionColumn < 0)
            {
                throw new InvalidOperationException(
                    "BOM 파일에서 '품목 설명' 또는 'Description' 열을 찾을 수 없습니다.");
            }

            var matchColumns = EnsureMatchColumns(outputTable, headerRowIndex);
            var matchedCount = 0;
            var unmatchedCount = 0;
            var unmatchedRows = new List<int>();
            var totalRows = Math.Max(0, outputTable.Rows.Count - (headerRowIndex + 1));
            var processedRows = 0;
            var dataStartRow = headerRowIndex + 1;
            var dataEndRow = headerRowIndex;

            using var repository = new BomDbRepository(databasePath);
            var allParts = repository.GetAllParts();

            for (var rowIndex = dataStartRow; rowIndex < outputTable.Rows.Count; rowIndex++)
            {
                processedRows++;
                onProgress?.Invoke(processedRows, totalRows);

                var row = outputTable.Rows[rowIndex];
                if (ShouldSkipRow(row, itemDescriptionColumn))
                {
                    continue;
                }

                dataEndRow = rowIndex;
                var itemDescription = GetCellText(row[itemDescriptionColumn]);
                var matchResult = BomDescriptionMatchEngine.Match(itemDescription, allParts);
                if (matchResult != null)
                {
                    row[matchColumns.PartNameColumn] = matchResult.Record.PartName;
                    row[matchColumns.PartNumberColumn] = matchResult.Record.ItemNumber;
                    row[matchColumns.ManufacturerColumn] = matchResult.Record.Manufacturer;
                    row[matchColumns.DbDescriptionColumn] =
                        BomDescriptionMatchEngine.FormatDbDescriptionForOutput(matchResult);
                    matchedCount++;
                }
                else
                {
                    row[matchColumns.PartNameColumn] = string.Empty;
                    row[matchColumns.PartNumberColumn] = UnregisteredPartText;
                    row[matchColumns.ManufacturerColumn] = UnregisteredPartText;
                    row[matchColumns.DbDescriptionColumn] = string.Empty;
                    unmatchedRows.Add(rowIndex);
                    unmatchedCount++;
                }
            }

            var totalProcessed = matchedCount + unmatchedCount;
            return new PdbBomDbMatchResult
            {
                OutputTable = outputTable,
                TotalProcessed = totalProcessed,
                MatchedCount = matchedCount,
                UnmatchedCount = unmatchedCount,
                Layout = new PdbBomDbMatchLayout
                {
                    HeaderRow = headerRowIndex,
                    DataStartRow = dataStartRow,
                    DataEndRow = dataEndRow,
                    PartNameColumn = matchColumns.PartNameColumn,
                    PartNumberColumn = matchColumns.PartNumberColumn,
                    ManufacturerColumn = matchColumns.ManufacturerColumn,
                    DbDescriptionColumn = matchColumns.DbDescriptionColumn,
                    UnmatchedRows = unmatchedRows
                }
            };
        }

        private static int FindItemDescriptionHeaderRow(DataTable table)
        {
            var searchLimit = Math.Min(table.Rows.Count, 100);
            for (var rowIndex = 0; rowIndex < searchLimit; rowIndex++)
            {
                if (FindItemDescriptionColumn(table.Rows[rowIndex]) >= 0)
                {
                    return rowIndex;
                }
            }

            return -1;
        }

        private static int FindItemDescriptionColumn(DataRow headerRow)
        {
            for (var columnIndex = 0; columnIndex < headerRow.Table.Columns.Count; columnIndex++)
            {
                if (IsItemDescriptionHeader(GetCellText(headerRow[columnIndex])))
                {
                    return columnIndex;
                }
            }

            return -1;
        }

        private static bool IsItemDescriptionHeader(string headerText)
        {
            if (string.IsNullOrWhiteSpace(headerText))
            {
                return false;
            }

            var normalized = headerText.Trim();
            if (normalized.Contains("DB", StringComparison.OrdinalIgnoreCase)
                && normalized.Contains("품목", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return ItemDescriptionHeaders.Any(candidate =>
                normalized.Equals(candidate, StringComparison.OrdinalIgnoreCase));
        }

        private static MatchColumnMap EnsureMatchColumns(DataTable table, int headerRowIndex)
        {
            var headerRow = table.Rows[headerRowIndex];
            return new MatchColumnMap
            {
                PartNameColumn = FindOrAddColumn(table, headerRow, PartNameHeaders, "품명"),
                PartNumberColumn = FindOrAddColumn(table, headerRow, PartNumberHeaders, "품목 번호"),
                ManufacturerColumn = FindOrAddColumn(table, headerRow, ManufacturerHeaders, "제조사"),
                DbDescriptionColumn = FindOrAddColumn(table, headerRow, DbDescriptionHeaders, "DB 품목 설명")
            };
        }

        private static int FindOrAddColumn(
            DataTable table,
            DataRow headerRow,
            IReadOnlyList<string> headerCandidates,
            string defaultHeaderName)
        {
            for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                var headerText = GetCellText(headerRow[columnIndex]);
                if (headerCandidates.Any(candidate =>
                        headerText.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
                {
                    return columnIndex;
                }
            }

            var newColumn = table.Columns.Add($"F{table.Columns.Count}", typeof(string));
            headerRow[newColumn] = defaultHeaderName;
            return newColumn.Ordinal;
        }

        private static bool ShouldSkipRow(DataRow row, int itemDescriptionColumn)
        {
            if (IsSeparatorRow(row))
            {
                return true;
            }

            var itemDescription = GetCellText(row[itemDescriptionColumn]);
            if (string.IsNullOrWhiteSpace(itemDescription))
            {
                return true;
            }

            if (itemDescription.Equals("NC List", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return IsSeparatorText(itemDescription);
        }

        private static bool IsSeparatorRow(DataRow row)
        {
            var hasContent = false;
            for (var columnIndex = 0; columnIndex < row.Table.Columns.Count; columnIndex++)
            {
                var text = GetCellText(row[columnIndex]);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                hasContent = true;
                if (!IsSeparatorText(text))
                {
                    return false;
                }
            }

            return hasContent;
        }

        private static bool IsSeparatorText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            var trimmed = text.Trim();
            if (SeparatorTextPattern.IsMatch(trimmed))
            {
                return true;
            }

            return trimmed.All(character => !char.IsLetterOrDigit(character));
        }

        private static string GetCellText(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            return Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
        }

        private sealed class MatchColumnMap
        {
            public int PartNameColumn { get; init; }
            public int PartNumberColumn { get; init; }
            public int ManufacturerColumn { get; init; }
            public int DbDescriptionColumn { get; init; }
        }
    }
}
