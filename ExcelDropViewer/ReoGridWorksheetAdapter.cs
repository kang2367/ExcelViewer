using System.Data;
using unvell.ReoGrid;

namespace ExcelDropViewer
{
    internal static class ReoGridWorksheetAdapter
    {
        private const string MaxTextLengthKey = "MaxTextLength";
        private const string IsResultColumnKey = "IsResultColumn";

        public static bool HasWorksheetData(Worksheet sheet)
        {
            return sheet.MaxContentRow >= 0 && sheet.MaxContentCol >= 0;
        }

        public static DataTable ToDataTable(Worksheet sheet)
        {
            var rowCount = Math.Max(0, sheet.MaxContentRow + 1);
            var columnCount = Math.Max(0, sheet.MaxContentCol + 1);

            if (rowCount == 0 || columnCount == 0)
            {
                return new DataTable();
            }

            var table = new DataTable();
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var column = table.Columns.Add($"F{columnIndex}", typeof(string));
                column.ExtendedProperties[MaxTextLengthKey] = 0;
            }

            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = table.NewRow();
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var text = GetCellText(sheet, rowIndex, columnIndex);
                    row[columnIndex] = text;
                    UpdateMaxTextLength(table.Columns[columnIndex], text);
                }

                table.Rows.Add(row);
            }

            return table;
        }

        public static void ApplyDataTable(Worksheet sheet, DataTable table)
        {
            sheet.Reset();

            if (table.Rows.Count == 0 || table.Columns.Count == 0)
            {
                return;
            }

            sheet.Resize(table.Rows.Count, table.Columns.Count);

            for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    var column = table.Columns[columnIndex];
                    var value = table.Rows[rowIndex][columnIndex];
                    var text = value == DBNull.Value ? string.Empty : System.Convert.ToString(value) ?? string.Empty;
                    sheet.SetCellData(rowIndex, columnIndex, text);

                    if (column.ExtendedProperties.ContainsKey(IsResultColumnKey)
                        && column.ExtendedProperties[IsResultColumnKey] is true
                        && string.Equals(text.Trim(), "NG", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.SetRangeStyles(
                            new RangePosition(rowIndex, columnIndex, 1, 1),
                            new WorksheetRangeStyle
                            {
                                Flag = PlainStyleFlag.BackColor,
                                BackColor = System.Drawing.Color.FromArgb(255, 228, 228)
                            });
                    }
                }
            }

            sheet.AutoFitColumnWidth(0, true);
        }

        public static int TryGetSelectedRowIndex(ReoGridControl grid)
        {
            var sheet = grid.CurrentWorksheet;
            if (sheet == null)
            {
                return -1;
            }

            var range = sheet.SelectionRange;
            return range.Row >= 0 ? range.Row : -1;
        }

        public static string GetCellText(Worksheet sheet, int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || columnIndex < 0)
            {
                return string.Empty;
            }

            return sheet.GetCellText(rowIndex, columnIndex)?.Trim() ?? string.Empty;
        }

        private static void UpdateMaxTextLength(DataColumn column, string value)
        {
            var currentMax = column.ExtendedProperties[MaxTextLengthKey] is int max ? max : 0;
            if (value.Length > currentMax)
            {
                column.ExtendedProperties[MaxTextLengthKey] = value.Length;
            }
        }
    }
}
