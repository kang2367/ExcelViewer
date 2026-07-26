using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;

namespace ExcelDropViewer
{
    internal static class BomDbImporter
    {
        private static readonly string[] PartNameHeaders = { "품명", "Part Name", "Item Name" };
        private static readonly string[] ItemNumberHeaders = { "품목 번호", "Part Number", "Part No" };
        private static readonly string[] ItemDescriptionHeaders = { "품목 설명", "Description" };
        private static readonly string[] ManufacturerHeaders = { "제조사", "Manufacturer", "Mfr" };
        private static readonly string[] OperatingTemperatureHeaders = { "동작 온도", "Operating Temperature", "Temp" };

        public static BomDbImportResult Import(
            DataTable table,
            int selectedHeaderRowIndex,
            Window ownerWindow,
            Action<int, int>? onProgress = null)
        {
            if (table == null || table.Rows.Count == 0)
            {
                throw new InvalidOperationException("저장할 데이터가 없습니다.");
            }

            var mapping = ResolveHeaderMapping(table, selectedHeaderRowIndex);
            var insertedCount = 0;
            var updatedCount = 0;
            var skippedCount = 0;
            var cancelled = false;
            var skipAllDuplicates = false;
            var totalRows = Math.Max(0, table.Rows.Count - (mapping.HeaderRowIndex + 1));

            var databasePath = BomDbRepository.GetDefaultDatabasePath();
            var backupFileName = BomDbBackupService.CreateBackupIfExists(databasePath);

            using var repository = new BomDbRepository(databasePath);
            repository.BeginTransaction();

            try
            {
                for (var rowIndex = mapping.HeaderRowIndex + 1; rowIndex < table.Rows.Count; rowIndex++)
                {
                    onProgress?.Invoke(rowIndex - mapping.HeaderRowIndex, totalRows);

                    var row = table.Rows[rowIndex];
                    var itemNumber = GetCellText(row[mapping.ItemNumberColumn]);
                    if (string.IsNullOrWhiteSpace(itemNumber))
                    {
                        continue;
                    }

                    var record = new BomPartRecord
                    {
                        PartName = GetCellText(row[mapping.PartNameColumn]),
                        ItemNumber = itemNumber,
                        ItemDescription = GetCellText(row[mapping.ItemDescriptionColumn]),
                        Manufacturer = GetCellText(row[mapping.ManufacturerColumn]),
                        OperatingTemperature = GetCellText(row[mapping.OperatingTemperatureColumn])
                    };

                    if (repository.ExistsByItemNumber(itemNumber))
                    {
                        var resolution = ResolveDuplicateAction(
                            ownerWindow,
                            itemNumber,
                            ref skipAllDuplicates);

                        switch (resolution)
                        {
                            case BomDuplicateResolution.Update:
                                repository.Update(record);
                                updatedCount++;
                                break;
                            case BomDuplicateResolution.Skip:
                                skippedCount++;
                                break;
                            case BomDuplicateResolution.Cancel:
                                cancelled = true;
                                break;
                        }
                    }
                    else
                    {
                        repository.Insert(record);
                        insertedCount++;
                    }

                    if (cancelled)
                    {
                        break;
                    }
                }

                if (cancelled)
                {
                    repository.RollbackTransaction();
                }
                else
                {
                    repository.CommitTransaction();
                }
            }
            catch
            {
                repository.RollbackTransaction();
                throw;
            }

            return new BomDbImportResult(
                insertedCount,
                updatedCount,
                skippedCount,
                cancelled,
                repository.DatabasePath,
                mapping.HeaderRowIndex,
                backupFileName);
        }

        private static BomDuplicateResolution ResolveDuplicateAction(
            Window ownerWindow,
            string itemNumber,
            ref bool skipAllDuplicates)
        {
            if (skipAllDuplicates)
            {
                return BomDuplicateResolution.Skip;
            }

            var resolution = BomDuplicateItemDialog.Show(ownerWindow, itemNumber);
            if (resolution == BomDuplicateResolution.AllSkip)
            {
                skipAllDuplicates = true;
                return BomDuplicateResolution.Skip;
            }

            return resolution;
        }

        private static BomDbColumnMapping ResolveHeaderMapping(DataTable table, int selectedHeaderRowIndex)
        {
            if (selectedHeaderRowIndex >= 0 && selectedHeaderRowIndex < table.Rows.Count)
            {
                var selectedMapping = TryMapColumns(table.Rows[selectedHeaderRowIndex], selectedHeaderRowIndex);
                if (selectedMapping != null)
                {
                    return selectedMapping;
                }
            }

            var candidateRows = new List<int>();
            if (table.Rows.Count > 0)
            {
                candidateRows.Add(0);
            }

            if (table.Rows.Count > 1)
            {
                candidateRows.Add(1);
            }

            foreach (var rowIndex in candidateRows)
            {
                if (rowIndex == selectedHeaderRowIndex)
                {
                    continue;
                }

                var mapping = TryMapColumns(table.Rows[rowIndex], rowIndex);
                if (mapping != null)
                {
                    return mapping;
                }
            }

            throw new InvalidOperationException(
                "헤더 행에서 '품명', '품목 번호', '품목 설명', '제조사', '동작 온도' 열을 찾을 수 없습니다. 헤더 행을 선택하거나 엑셀 헤더 명칭을 확인해 주세요.");
        }

        private static BomDbColumnMapping? TryMapColumns(DataRow headerRow, int headerRowIndex)
        {
            var usedColumns = new HashSet<int>();

            var partNameColumn = FindColumnIndexContaining(headerRow, PartNameHeaders, usedColumns);
            var itemNumberColumn = FindColumnIndexContaining(headerRow, ItemNumberHeaders, usedColumns);
            var itemDescriptionColumn = FindColumnIndexContaining(headerRow, ItemDescriptionHeaders, usedColumns);
            var manufacturerColumn = FindColumnIndexContaining(headerRow, ManufacturerHeaders, usedColumns);
            var operatingTemperatureColumn = FindColumnIndexContaining(headerRow, OperatingTemperatureHeaders, usedColumns);

            if (partNameColumn < 0
                || itemNumberColumn < 0
                || itemDescriptionColumn < 0
                || manufacturerColumn < 0
                || operatingTemperatureColumn < 0)
            {
                return null;
            }

            return new BomDbColumnMapping(
                headerRowIndex,
                partNameColumn,
                itemNumberColumn,
                itemDescriptionColumn,
                manufacturerColumn,
                operatingTemperatureColumn);
        }

        private static int FindColumnIndexContaining(
            DataRow headerRow,
            IReadOnlyList<string> headerKeywords,
            ISet<int> usedColumns)
        {
            foreach (var keyword in OrderKeywordsBySpecificity(headerKeywords))
            {
                for (var columnIndex = 0; columnIndex < headerRow.Table.Columns.Count; columnIndex++)
                {
                    if (usedColumns.Contains(columnIndex))
                    {
                        continue;
                    }

                    var headerText = GetCellText(headerRow[columnIndex]);
                    if (headerText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        usedColumns.Add(columnIndex);
                        return columnIndex;
                    }
                }
            }

            return -1;
        }

        private static IEnumerable<string> OrderKeywordsBySpecificity(IReadOnlyList<string> headerKeywords)
        {
            return headerKeywords
                .OrderByDescending(static keyword => keyword.Length)
                .ThenBy(static keyword => keyword, StringComparer.OrdinalIgnoreCase);
        }

        private static string GetCellText(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            return System.Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
        }
    }

    internal sealed class BomDbColumnMapping
    {
        public BomDbColumnMapping(
            int headerRowIndex,
            int partNameColumn,
            int itemNumberColumn,
            int itemDescriptionColumn,
            int manufacturerColumn,
            int operatingTemperatureColumn)
        {
            HeaderRowIndex = headerRowIndex;
            PartNameColumn = partNameColumn;
            ItemNumberColumn = itemNumberColumn;
            ItemDescriptionColumn = itemDescriptionColumn;
            ManufacturerColumn = manufacturerColumn;
            OperatingTemperatureColumn = operatingTemperatureColumn;
        }

        public int HeaderRowIndex { get; }

        public int PartNameColumn { get; }

        public int ItemNumberColumn { get; }

        public int ItemDescriptionColumn { get; }

        public int ManufacturerColumn { get; }

        public int OperatingTemperatureColumn { get; }
    }

    internal sealed class BomDbImportResult
    {
        public BomDbImportResult(
            int insertedCount,
            int updatedCount,
            int skippedCount,
            bool cancelled,
            string databasePath,
            int headerRowIndex,
            string? backupFileName)
        {
            InsertedCount = insertedCount;
            UpdatedCount = updatedCount;
            SkippedCount = skippedCount;
            Cancelled = cancelled;
            DatabasePath = databasePath;
            HeaderRowIndex = headerRowIndex;
            BackupFileName = backupFileName;
        }

        public int InsertedCount { get; }

        public int UpdatedCount { get; }

        public int SkippedCount { get; }

        public bool Cancelled { get; }

        public string DatabasePath { get; }

        public int HeaderRowIndex { get; }

        public string? BackupFileName { get; }

        public string BuildSummaryMessage()
        {
            if (Cancelled)
            {
                return $"BOM DB 업데이트가 취소되었습니다.\n" +
                       $"- 백업 파일: {FormatBackupFileName()}\n" +
                       $"- 신규 추가: {InsertedCount}건, 업데이트: {UpdatedCount}건, 건너뜀(Skip): {SkippedCount}건\n" +
                       "작업이 취소되어 변경 사항이 저장되지 않았습니다.";
            }

            return $"BOM DB 업데이트가 완료되었습니다.\n" +
                   $"- 백업 파일: {FormatBackupFileName()}\n" +
                   $"- 신규 추가: {InsertedCount}건, 업데이트: {UpdatedCount}건, 건너뜀(Skip): {SkippedCount}건";
        }

        private string FormatBackupFileName()
        {
            return string.IsNullOrWhiteSpace(BackupFileName) ? "없음" : BackupFileName;
        }
    }
}
