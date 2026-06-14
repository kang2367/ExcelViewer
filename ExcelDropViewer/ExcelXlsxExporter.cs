using System.Data;
using System.Globalization;
using ClosedXML.Excel;

namespace ExcelDropViewer
{
    internal static class ExcelXlsxExporter
    {
        public static void SaveDataTable(DataTable table, string filePath)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet1");
            var usedRange = worksheet.Range(
                1,
                1,
                Math.Max(1, table.Rows.Count),
                Math.Max(1, table.Columns.Count));
            usedRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            usedRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var dataRow = table.Rows[rowIndex];
                for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    var cell = worksheet.Cell(rowIndex + 1, columnIndex + 1);
                    var text = GetCellText(dataRow[columnIndex]);
                    cell.Value = text;
                    cell.Style.Alignment.WrapText = text.Contains('\n', StringComparison.Ordinal);
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                }
            }

            worksheet.Columns().AdjustToContents(1, 50);
            workbook.SaveAs(filePath);
        }

        private static string GetCellText(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            return Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty;
        }
    }
}
