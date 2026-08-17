using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelDropViewer
{
    internal sealed class PbaBomTransformResult
    {
        public DataTable OutputTable { get; init; } = new();
        public int NormalItemCount { get; init; }
        public int NcItemCount { get; init; }
        public PbaBomOutputLayout Layout { get; init; } = new();
    }

    internal sealed class PbaBomOutputLayout
    {
        public int TemplateHeaderRow { get; init; }
        public int DataStartRow { get; init; }
        public int DataEndRow { get; init; }
        public int ReferenceColumn { get; init; } = -1;
        public int ItemDescriptionColumn { get; init; } = -1;
        public int NcListTitleRow { get; init; } = -1;
        public int NcListTitleColumn { get; init; }
    }

    internal static class PbaBomTransformer
    {
        private static readonly string[] RawRequiredHeaders = { "Item", "Part", "Quantity", "Reference" };
        private static readonly string[] TemplateRequiredHeaders = { "No", "Reference", "품명", "수량" };

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

        private static readonly Regex MountingHoleOrTestPointPattern = new(
            @"\b(MH|TP)\d+\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex CapacitorPattern = new(
            @"\d+(?:\.\d+)?\s*(?:uF|nF|pF|µF|uf|nf|pf)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ResistorValuePattern = new(
            @"^\d+(?:\.\d+)?[KM]?$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex WattagePattern = new(
            @"\d+/\d+W|\d+(?:\.\d+)?W",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex UnderscoreWattageResistorPattern = new(
            @"^\d+(?:\.\d+)?[KM]?_(?:\d+(?:\.\d+)?(?:/\d+)?W)$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex NoConnectPartPattern = new(
            @"^(?:.+/NC|NC)$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex PureNumericPartPattern = new(
            @"^\d+(?:\.\d+)?$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex ConnectorReferencePattern = new(
            @"^J\d+",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex InductorReferencePattern = new(
            @"^L\d+",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex SeparatorTextPattern = new(
            @"^[-_=~*.·•─━┅┄┈┉―﹣－\s]+$",
            RegexOptions.Compiled);

        public static PbaBomTransformResult Transform(
            DataTable rawTable,
            DataTable templateTable,
            Action<int, int>? onProgress = null)
        {
            if (rawTable.Rows.Count == 0)
            {
                throw new InvalidOperationException("Raw Data에 변환할 데이터가 없습니다.");
            }

            if (templateTable.Rows.Count == 0)
            {
                throw new InvalidOperationException("Template에 변환할 데이터가 없습니다.");
            }

            var rawHeaderRow = FindHeaderRow(rawTable, RawRequiredHeaders);
            if (rawHeaderRow < 0)
            {
                throw new InvalidOperationException(
                    "Raw Data에서 Item, Part, Quantity, Reference 헤더 행을 찾을 수 없습니다.");
            }

            var templateHeaderRow = FindHeaderRow(templateTable, TemplateRequiredHeaders);
            if (templateHeaderRow < 0)
            {
                throw new InvalidOperationException(
                    "Template에서 No, Reference, 품명, 수량 헤더 행을 찾을 수 없습니다.");
            }

            var rawColumns = ResolveRawColumns(rawTable.Rows[rawHeaderRow]);
            ValidateRawColumns(rawColumns);
            var templateColumns = ResolveTemplateColumns(templateTable.Rows[templateHeaderRow]);
            ValidateTemplateColumns(templateColumns);

            var normalItems = new List<PbaBomRowData>();
            var ncItems = new List<PbaBomRowData>();
            var totalRows = Math.Max(0, rawTable.Rows.Count - (rawHeaderRow + 1));
            var processedRows = 0;
            var dataReadingStarted = false;

            for (var rowIndex = rawHeaderRow + 1; rowIndex < rawTable.Rows.Count; rowIndex++)
            {
                processedRows++;
                onProgress?.Invoke(processedRows, totalRows);

                var sourceRow = rawTable.Rows[rowIndex];
                if (IsSeparatorRow(sourceRow))
                {
                    continue;
                }

                var itemValue = GetCellValue(sourceRow, rawColumns.Item);
                if (!dataReadingStarted)
                {
                    if (!IsNumericItem(itemValue))
                    {
                        continue;
                    }

                    dataReadingStarted = true;
                }

                if (IsEmptyDataRow(sourceRow, rawColumns))
                {
                    continue;
                }

                var reference = GetCellValue(sourceRow, rawColumns.Reference);
                if (ContainsMountingHoleOrTestPoint(reference))
                {
                    continue;
                }

                var part = GetCellValue(sourceRow, rawColumns.Part);
                var rowData = BuildRowData(sourceRow, rawColumns);

                if (ContainsNoConnect(part))
                {
                    ncItems.Add(rowData);
                }
                else
                {
                    normalItems.Add(rowData);
                }
            }

            var outputTable = BuildOutputTable(
                templateTable,
                templateHeaderRow,
                templateColumns,
                normalItems,
                ncItems,
                out var layout);

            return new PbaBomTransformResult
            {
                OutputTable = outputTable,
                NormalItemCount = normalItems.Count,
                NcItemCount = ncItems.Count,
                Layout = layout
            };
        }

        private static int FindHeaderRow(DataTable table, IReadOnlyList<string> requiredHeaders)
        {
            var searchLimit = Math.Min(table.Rows.Count, 100);
            for (var rowIndex = 0; rowIndex < searchLimit; rowIndex++)
            {
                if (RowContainsAllHeaders(table.Rows[rowIndex], requiredHeaders))
                {
                    return rowIndex;
                }
            }

            return -1;
        }

        private static bool RowContainsAllHeaders(DataRow row, IReadOnlyList<string> requiredHeaders)
        {
            var cellTexts = row.ItemArray.Select(value => GetCellText(value)).ToList();
            return requiredHeaders.All(required =>
                cellTexts.Any(cell => HeaderMatches(cell, required)));
        }

        private static bool HeaderMatches(string cellText, string requiredHeader)
        {
            if (string.IsNullOrWhiteSpace(cellText))
            {
                return false;
            }

            var normalized = cellText.Trim();
            if (requiredHeader.Equals("No", StringComparison.OrdinalIgnoreCase))
            {
                return normalized.Equals("No", StringComparison.OrdinalIgnoreCase)
                    || normalized.Equals("NO", StringComparison.OrdinalIgnoreCase)
                    || normalized.Equals("No.", StringComparison.OrdinalIgnoreCase);
            }

            return normalized.Equals(requiredHeader, StringComparison.OrdinalIgnoreCase)
                || normalized.Contains(requiredHeader, StringComparison.OrdinalIgnoreCase);
        }

        private static RawColumnMap ResolveRawColumns(DataRow headerRow)
        {
            return new RawColumnMap
            {
                Item = FindColumnIndex(headerRow, "Item", "아이템"),
                Part = FindColumnIndex(headerRow, "Part", "Part Number", "부품번호"),
                Quantity = FindColumnIndex(headerRow, "Quantity", "Q'ty", "Qty", "QTY"),
                Reference = FindReferenceColumn(headerRow),
                Size = FindColumnIndex(headerRow, "Size", "Package", "Footprint"),
                Voltage = FindColumnIndex(headerRow, "Voltage", "Volt", "Rated Voltage"),
                Tolerance = FindColumnIndex(headerRow, "Tolerance", "Tol"),
                Spec = FindColumnIndex(headerRow, "Spec", "Specification", "Specification/Spec"),
                Description = FindColumnIndex(headerRow, "Description", "Desc", "Value", "Comment")
            };
        }

        private static TemplateColumnMap ResolveTemplateColumns(DataRow headerRow)
        {
            return new TemplateColumnMap
            {
                No = FindColumnIndex(headerRow, "No", "NO", "No."),
                Reference = FindReferenceColumn(headerRow),
                PartName = FindColumnIndex(headerRow, "품명"),
                Quantity = FindColumnIndex(headerRow, "수량", "Quantity", "Q'ty", "Qty", "QTY"),
                PartNumber = FindColumnIndex(headerRow, "품목번호", "Part", "품번", "Part Number"),
                ItemDescription = FindColumnIndex(headerRow, "품목 설명", "Description", "품목설명")
            };
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
            return ReferenceHeaderNames.Any(candidate =>
                normalized.Equals(candidate, StringComparison.OrdinalIgnoreCase));
        }

        private static int FindColumnIndex(DataRow headerRow, params string[] candidates)
        {
            for (var columnIndex = 0; columnIndex < headerRow.Table.Columns.Count; columnIndex++)
            {
                var headerText = GetCellText(headerRow[columnIndex]);
                if (candidates.Any(candidate => HeaderMatches(headerText, candidate)))
                {
                    return columnIndex;
                }
            }

            return -1;
        }

        private static void ValidateRawColumns(RawColumnMap columns)
        {
            if (columns.Item < 0 || columns.Part < 0 || columns.Quantity < 0 || columns.Reference < 0)
            {
                throw new InvalidOperationException(
                    "Raw Data 헤더 행에서 Item, Part, Quantity, Reference 열을 모두 찾을 수 없습니다.");
            }
        }

        private static void ValidateTemplateColumns(TemplateColumnMap columns)
        {
            if (columns.No < 0 || columns.Reference < 0 || columns.PartName < 0 || columns.Quantity < 0)
            {
                throw new InvalidOperationException(
                    "Template 헤더 행에서 No, Reference, 품명, 수량 열을 모두 찾을 수 없습니다.");
            }
        }

        private static bool IsNumericItem(string itemValue)
        {
            if (string.IsNullOrWhiteSpace(itemValue))
            {
                return false;
            }

            return int.TryParse(itemValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out _)
                || double.TryParse(itemValue, NumberStyles.Number, CultureInfo.InvariantCulture, out _);
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

        private static bool IsEmptyDataRow(DataRow row, RawColumnMap columns)
        {
            return string.IsNullOrWhiteSpace(GetCellValue(row, columns.Item))
                && string.IsNullOrWhiteSpace(GetCellValue(row, columns.Part))
                && string.IsNullOrWhiteSpace(GetCellValue(row, columns.Reference))
                && string.IsNullOrWhiteSpace(GetCellValue(row, columns.Quantity));
        }

        private static bool ContainsMountingHoleOrTestPoint(string reference)
        {
            return !string.IsNullOrWhiteSpace(reference)
                && MountingHoleOrTestPointPattern.IsMatch(reference);
        }

        private static bool ContainsNoConnect(string part)
        {
            return !string.IsNullOrWhiteSpace(part)
                && NoConnectPartPattern.IsMatch(part.Trim());
        }

        private static PbaBomRowData BuildRowData(DataRow sourceRow, RawColumnMap columns)
        {
            var part = GetCellValue(sourceRow, columns.Part);
            var reference = GetCellValue(sourceRow, columns.Reference);
            var size = GetCellValue(sourceRow, columns.Size);
            var voltage = GetCellValue(sourceRow, columns.Voltage);
            var tolerance = GetCellValue(sourceRow, columns.Tolerance);
            var spec = GetCellValue(sourceRow, columns.Spec);
            var description = GetCellValue(sourceRow, columns.Description);

            return new PbaBomRowData
            {
                No = GetCellValue(sourceRow, columns.Item),
                Reference = reference,
                Quantity = GetCellValue(sourceRow, columns.Quantity),
                PartNumber = FormatPartNumberForOutput(part),
                ItemDescription = BuildItemDescription(
                    part,
                    reference,
                    size,
                    voltage,
                    tolerance,
                    spec,
                    description)
            };
        }

        internal static string FormatPartNumberForOutput(string part)
        {
            if (string.IsNullOrWhiteSpace(part))
            {
                return part;
            }

            var trimmed = part.Trim();
            if (!IsUnderscoreWattageResistorPart(trimmed))
            {
                return trimmed;
            }

            var underscoreIndex = trimmed.IndexOf('_');
            return trimmed[..underscoreIndex];
        }

        private static string GetResistorPartValueForDescription(string part)
        {
            return FormatPartNumberForOutput(part);
        }

        private static string JoinDescriptionParts(params string?[] parts)
        {
            return string.Join(
                ", ",
                parts.Where(part => !string.IsNullOrWhiteSpace(part)).Select(part => part!.Trim()));
        }

        internal static string BuildItemDescription(
            string part,
            string reference,
            string size,
            string voltage,
            string tolerance,
            string spec,
            string fallbackDescription)
        {
            if (IsCapacitor(part))
            {
                return JoinDescriptionParts("CAP", "MLCC", size, part, voltage, tolerance);
            }

            if (IsPureNumericPart(part))
            {
                var primaryReference = GetPrimaryReference(reference);
                if (IsConnectorReference(primaryReference))
                {
                    return JoinDescriptionParts("CONN", part);
                }

                if (IsInductorReference(primaryReference))
                {
                    var toleranceOrSpec = !string.IsNullOrWhiteSpace(tolerance) ? tolerance : spec;
                    return JoinDescriptionParts("IND", size, part, toleranceOrSpec);
                }
            }

            if (IsResistor(part))
            {
                var partForDescription = GetResistorPartValueForDescription(part);
                var wattage = ExtractWattage(part);
                return JoinDescriptionParts("RES", size, partForDescription, tolerance, wattage);
            }

            return !string.IsNullOrWhiteSpace(fallbackDescription) ? fallbackDescription : part;
        }

        private static bool IsPureNumericPart(string part)
        {
            return !string.IsNullOrWhiteSpace(part) && PureNumericPartPattern.IsMatch(part.Trim());
        }

        private static string GetPrimaryReference(string reference)
        {
            if (string.IsNullOrWhiteSpace(reference))
            {
                return string.Empty;
            }

            var token = reference
                .Split(new[] { ',', ' ', ';', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .FirstOrDefault();

            return token?.Trim() ?? string.Empty;
        }

        private static bool IsConnectorReference(string reference)
        {
            return !string.IsNullOrWhiteSpace(reference)
                && ConnectorReferencePattern.IsMatch(reference.Trim());
        }

        private static bool IsInductorReference(string reference)
        {
            return !string.IsNullOrWhiteSpace(reference)
                && InductorReferencePattern.IsMatch(reference.Trim());
        }

        private static bool IsCapacitor(string part)
        {
            return !string.IsNullOrWhiteSpace(part) && CapacitorPattern.IsMatch(part);
        }

        private static bool IsResistor(string part)
        {
            if (string.IsNullOrWhiteSpace(part))
            {
                return false;
            }

            var trimmed = part.Trim();
            if (IsUnderscoreWattageResistorPart(trimmed))
            {
                return true;
            }

            var normalized = RemoveWattage(trimmed).Trim();
            return ResistorValuePattern.IsMatch(normalized);
        }

        private static bool IsUnderscoreWattageResistorPart(string part)
        {
            return UnderscoreWattageResistorPattern.IsMatch(part.Trim());
        }

        private static string RemoveWattage(string part)
        {
            return WattagePattern.Replace(part, string.Empty).Trim();
        }

        private static string ExtractWattage(string part)
        {
            var underscoreIndex = part.IndexOf('_');
            if (underscoreIndex >= 0)
            {
                var afterUnderscore = part[(underscoreIndex + 1)..];
                var underscoreMatch = WattagePattern.Match(afterUnderscore);
                if (underscoreMatch.Success)
                {
                    return underscoreMatch.Value;
                }
            }

            var match = WattagePattern.Match(part);
            return match.Success ? match.Value : string.Empty;
        }

        private static DataTable BuildOutputTable(
            DataTable templateTable,
            int templateHeaderRow,
            TemplateColumnMap templateColumns,
            IReadOnlyList<PbaBomRowData> normalItems,
            IReadOnlyList<PbaBomRowData> ncItems,
            out PbaBomOutputLayout layout)
        {
            var outputTable = templateTable.Clone();
            for (var rowIndex = 0; rowIndex <= templateHeaderRow; rowIndex++)
            {
                CopyRow(templateTable.Rows[rowIndex], outputTable);
            }

            var writeRowIndex = templateHeaderRow + 1;
            var dataStartRow = writeRowIndex;
            var ncListTitleRow = -1;
            var ncListTitleColumn = templateColumns.No >= 0 ? templateColumns.No : 0;

            foreach (var item in normalItems)
            {
                WriteRow(outputTable, writeRowIndex++, templateColumns, item);
            }

            if (ncItems.Count > 0)
            {
                EnsureEmptyRow(outputTable, writeRowIndex++);
                ncListTitleRow = writeRowIndex++;
                WriteNcListTitleRow(outputTable, ncListTitleRow, ncListTitleColumn);

                foreach (var item in ncItems)
                {
                    WriteRow(outputTable, writeRowIndex++, templateColumns, item);
                }
            }

            layout = new PbaBomOutputLayout
            {
                TemplateHeaderRow = templateHeaderRow,
                DataStartRow = dataStartRow,
                DataEndRow = Math.Max(dataStartRow, writeRowIndex - 1),
                ReferenceColumn = templateColumns.Reference,
                ItemDescriptionColumn = templateColumns.ItemDescription,
                NcListTitleRow = ncListTitleRow,
                NcListTitleColumn = ncListTitleColumn
            };

            return outputTable;
        }

        private static void WriteNcListTitleRow(DataTable table, int rowIndex, int columnIndex)
        {
            EnsureRowExists(table, rowIndex);
            var row = table.Rows[rowIndex];
            for (var index = 0; index < table.Columns.Count; index++)
            {
                row[index] = string.Empty;
            }

            if (columnIndex >= 0)
            {
                row[columnIndex] = "NC List";
            }
        }

        private static void CopyRow(DataRow sourceRow, DataTable targetTable)
        {
            var targetRow = targetTable.NewRow();
            for (var columnIndex = 0; columnIndex < sourceRow.Table.Columns.Count; columnIndex++)
            {
                targetRow[columnIndex] = GetCellText(sourceRow[columnIndex]);
            }

            targetTable.Rows.Add(targetRow);
        }

        private static void EnsureEmptyRow(DataTable table, int rowIndex)
        {
            EnsureRowExists(table, rowIndex);
            var row = table.Rows[rowIndex];
            for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
            {
                row[columnIndex] = string.Empty;
            }
        }

        private static void WriteRow(
            DataTable table,
            int rowIndex,
            TemplateColumnMap columns,
            PbaBomRowData item)
        {
            EnsureRowExists(table, rowIndex);
            var row = table.Rows[rowIndex];

            SetCell(row, columns.No, item.No);
            SetCell(row, columns.Reference, item.Reference);
            SetCell(row, columns.Quantity, item.Quantity);
            SetCell(row, columns.PartNumber, item.PartNumber);
            SetCell(row, columns.ItemDescription, item.ItemDescription);
        }

        private static void EnsureRowExists(DataTable table, int rowIndex)
        {
            while (table.Rows.Count <= rowIndex)
            {
                table.Rows.Add(table.NewRow());
            }
        }

        private static void SetCell(DataRow row, int columnIndex, string value)
        {
            if (columnIndex < 0)
            {
                return;
            }

            row[columnIndex] = value;
        }

        private static string GetCellValue(DataRow row, int columnIndex)
        {
            if (columnIndex < 0)
            {
                return string.Empty;
            }

            return GetCellText(row[columnIndex]);
        }

        private static string GetCellText(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            return Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
        }

        private sealed class RawColumnMap
        {
            public int Item { get; init; } = -1;
            public int Part { get; init; } = -1;
            public int Quantity { get; init; } = -1;
            public int Reference { get; init; } = -1;
            public int Size { get; init; } = -1;
            public int Voltage { get; init; } = -1;
            public int Tolerance { get; init; } = -1;
            public int Spec { get; init; } = -1;
            public int Description { get; init; } = -1;
        }

        private sealed class TemplateColumnMap
        {
            public int No { get; init; } = -1;
            public int Reference { get; init; } = -1;
            public int PartName { get; init; } = -1;
            public int Quantity { get; init; } = -1;
            public int PartNumber { get; init; } = -1;
            public int ItemDescription { get; init; } = -1;
        }

        private sealed class PbaBomRowData
        {
            public string No { get; init; } = string.Empty;
            public string Reference { get; init; } = string.Empty;
            public string Quantity { get; init; } = string.Empty;
            public string PartNumber { get; init; } = string.Empty;
            public string ItemDescription { get; init; } = string.Empty;
        }
    }
}
