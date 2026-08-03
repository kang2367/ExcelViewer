using unvell.ReoGrid;

namespace ExcelDropViewer
{
    internal sealed class DigiKeySearchGridEntry
    {
        public string PartNumber { get; init; } = string.Empty;
        public string Manufacturer { get; init; } = string.Empty;
        public string Stock { get; init; } = string.Empty;
        public string UnitPrice { get; init; } = string.Empty;
        public IReadOnlySet<string> MatchKeys { get; init; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    internal static class DigiKeySearchResultGridMapper
    {
        public static DigiKeySearchGridEntry ToGridEntry(DigiKeyProductSummary summary)
        {
            var matchKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddMatchKey(matchKeys, summary.DisplayPartNumber);
            AddMatchKey(matchKeys, summary.SearchedPartNumber);
            AddMatchKey(matchKeys, summary.DigiKeyPartNumber);

            return new DigiKeySearchGridEntry
            {
                PartNumber = summary.DisplayPartNumber,
                Manufacturer = summary.DisplayManufacturer,
                Stock = summary.FormattedStock,
                UnitPrice = summary.FormattedBaseUnitPrice,
                MatchKeys = matchKeys
            };
        }

        private static void AddMatchKey(ISet<string> matchKeys, string? value)
        {
            if (string.IsNullOrWhiteSpace(value) || value == "-")
            {
                return;
            }

            matchKeys.Add(value.Trim());
        }
    }

    internal static class DigiKeySearchResultGridWriter
    {
        private static readonly string[] DefaultHeaders = { "품번", "제조사", "재고", "단가" };

        private static readonly string[] PartNumberHeaders =
        {
            "품번", "Part Number", "PartNumber", "부품번호", "부품 번호", "Manufacturer Part Number"
        };

        private static readonly string[] ManufacturerHeaders = { "제조사", "Manufacturer" };
        private static readonly string[] StockHeaders = { "재고", "Stock", "Quantity", "QuantityAvailable" };
        private static readonly string[] UnitPriceHeaders = { "단가", "Unit Price", "UnitPrice", "Price" };

        public static void Upsert(Worksheet sheet, DigiKeyProductSummary summary)
        {
            var entry = DigiKeySearchResultGridMapper.ToGridEntry(summary);

            if (!ReoGridWorksheetAdapter.HasWorksheetData(sheet))
            {
                InitializeGrid(sheet, entry);
                return;
            }

            var columnMap = ResolveColumnMap(sheet) ?? AppendColumnMap(sheet);
            var dataStartRow = 1;
            var existingRow = FindMatchingRow(sheet, columnMap.PartNumberColumn, entry.MatchKeys, dataStartRow);
            var targetRow = existingRow >= 0 ? existingRow : sheet.MaxContentRow + 1;

            if (targetRow < dataStartRow)
            {
                targetRow = dataStartRow;
            }

            EnsureWorksheetSize(sheet, targetRow, columnMap);
            WriteRow(sheet, targetRow, columnMap, entry);
            sheet.AutoFitColumnWidth(0, true);
        }

        private static void InitializeGrid(Worksheet sheet, DigiKeySearchGridEntry entry)
        {
            sheet.Reset();
            sheet.Resize(2, DefaultHeaders.Length);

            for (var columnIndex = 0; columnIndex < DefaultHeaders.Length; columnIndex++)
            {
                sheet.SetCellData(0, columnIndex, DefaultHeaders[columnIndex]);
            }

            WriteRow(sheet, 1, CreateDefaultColumnMap(), entry);
            sheet.AutoFitColumnWidth(0, true);
        }

        private static DigiKeyGridColumnMap AppendColumnMap(Worksheet sheet)
        {
            var startColumn = Math.Max(0, sheet.MaxContentCol + 1);
            EnsureWorksheetSize(sheet, 0, startColumn + DefaultHeaders.Length - 1);

            for (var columnIndex = 0; columnIndex < DefaultHeaders.Length; columnIndex++)
            {
                sheet.SetCellData(0, startColumn + columnIndex, DefaultHeaders[columnIndex]);
            }

            return new DigiKeyGridColumnMap(
                startColumn,
                startColumn + 1,
                startColumn + 2,
                startColumn + 3);
        }

        private static DigiKeyGridColumnMap? ResolveColumnMap(Worksheet sheet)
        {
            var partNumberColumn = FindColumnIndex(sheet, PartNumberHeaders);
            var manufacturerColumn = FindColumnIndex(sheet, ManufacturerHeaders);
            var stockColumn = FindColumnIndex(sheet, StockHeaders);
            var unitPriceColumn = FindColumnIndex(sheet, UnitPriceHeaders);

            if (partNumberColumn < 0 || manufacturerColumn < 0 || stockColumn < 0 || unitPriceColumn < 0)
            {
                return null;
            }

            return new DigiKeyGridColumnMap(
                partNumberColumn,
                manufacturerColumn,
                stockColumn,
                unitPriceColumn);
        }

        private static int FindColumnIndex(Worksheet sheet, IReadOnlyList<string> headerCandidates)
        {
            var columnCount = Math.Max(sheet.MaxContentCol + 1, DefaultHeaders.Length);
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var headerText = ReoGridWorksheetAdapter.GetCellText(sheet, 0, columnIndex);
                if (headerCandidates.Any(candidate =>
                        headerText.Equals(candidate, StringComparison.OrdinalIgnoreCase)))
                {
                    return columnIndex;
                }
            }

            return -1;
        }

        private static int FindMatchingRow(
            Worksheet sheet,
            int partNumberColumn,
            IReadOnlySet<string> matchKeys,
            int startRow)
        {
            for (var rowIndex = startRow; rowIndex <= sheet.MaxContentRow; rowIndex++)
            {
                var cellValue = ReoGridWorksheetAdapter.GetCellText(sheet, rowIndex, partNumberColumn);
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    continue;
                }

                if (matchKeys.Contains(cellValue.Trim()))
                {
                    return rowIndex;
                }
            }

            return -1;
        }

        private static void EnsureWorksheetSize(Worksheet sheet, int rowIndex, DigiKeyGridColumnMap columnMap)
        {
            var requiredRows = Math.Max(sheet.RowCount, rowIndex + 1);
            var requiredColumns = Math.Max(sheet.ColumnCount, columnMap.MaxColumnIndex + 1);
            sheet.Resize(requiredRows, requiredColumns);
        }

        private static void EnsureWorksheetSize(Worksheet sheet, int headerRowIndex, int lastColumnIndex)
        {
            var requiredRows = Math.Max(sheet.RowCount, headerRowIndex + 1);
            var requiredColumns = Math.Max(sheet.ColumnCount, lastColumnIndex + 1);
            sheet.Resize(requiredRows, requiredColumns);
        }

        private static void WriteRow(
            Worksheet sheet,
            int rowIndex,
            DigiKeyGridColumnMap columnMap,
            DigiKeySearchGridEntry entry)
        {
            sheet.SetCellData(rowIndex, columnMap.PartNumberColumn, entry.PartNumber);
            sheet.SetCellData(rowIndex, columnMap.ManufacturerColumn, entry.Manufacturer);
            sheet.SetCellData(rowIndex, columnMap.StockColumn, entry.Stock);
            sheet.SetCellData(rowIndex, columnMap.UnitPriceColumn, entry.UnitPrice);
        }

        private static DigiKeyGridColumnMap CreateDefaultColumnMap()
        {
            return new DigiKeyGridColumnMap(0, 1, 2, 3);
        }

        private sealed class DigiKeyGridColumnMap
        {
            public DigiKeyGridColumnMap(
                int partNumberColumn,
                int manufacturerColumn,
                int stockColumn,
                int unitPriceColumn)
            {
                PartNumberColumn = partNumberColumn;
                ManufacturerColumn = manufacturerColumn;
                StockColumn = stockColumn;
                UnitPriceColumn = unitPriceColumn;
            }

            public int PartNumberColumn { get; }
            public int ManufacturerColumn { get; }
            public int StockColumn { get; }
            public int UnitPriceColumn { get; }

            public int MaxColumnIndex => Math.Max(
                PartNumberColumn,
                Math.Max(ManufacturerColumn, Math.Max(StockColumn, UnitPriceColumn)));
        }
    }
}
