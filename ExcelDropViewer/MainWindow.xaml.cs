using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using ExcelDataReader;

namespace ExcelDropViewer
{
    public partial class MainWindow : Window
    {
        private const string DisplayHeaderKey = "DisplayHeader";
        private const string MaxTextLengthKey = "MaxTextLength";
        private const int WrapCharacterThreshold = 40;
        private const double WrapColumnWidth = 300;
        private const double WrapColumnMaxWidth = 350;

        public MainWindow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
        }

        private void LeftZone_DragOver(object sender, DragEventArgs e)
        {
            HandleZoneDragOver(e);
        }

        private void RightZone_DragOver(object sender, DragEventArgs e)
        {
            HandleZoneDragOver(e);
        }

        private async void LeftZone_Drop(object sender, DragEventArgs e)
        {
            var excelPath = TryGetDroppedExcelPath(e, out var hasFileDrop);
            if (excelPath == null)
            {
                if (hasFileDrop)
                {
                    ShowUnsupportedFileMessage();
                }

                return;
            }

            await LoadExcelAsync(excelPath, LeftExcelGrid, LeftDropHint);
        }

        private async void RightZone_Drop(object sender, DragEventArgs e)
        {
            var excelPath = TryGetDroppedExcelPath(e, out var hasFileDrop);
            if (excelPath == null)
            {
                if (hasFileDrop)
                {
                    ShowUnsupportedFileMessage();
                }

                return;
            }

            await LoadExcelAsync(excelPath, RightExcelGrid, RightDropHint);
        }

        private static void HandleZoneDragOver(DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            e.Effects = paths != null && paths.Any(IsSupportedExcelFile)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
            e.Handled = true;
        }

        /// <summary>
        /// 지정한 영역의 DataGrid에만 데이터를 바인딩합니다. 반대쪽 영역은 변경하지 않습니다.
        /// </summary>
        private async Task LoadExcelAsync(string filePath, DataGrid targetGrid, TextBlock dropHint)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                var table = await Task.Run(() => LoadFirstSheet(filePath));

                BindDataGrid(table, targetGrid);
                dropHint.Visibility = Visibility.Collapsed;
                targetGrid.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"파일을 읽는 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private static string TryGetDroppedExcelPath(DragEventArgs e, out bool hasFileDrop)
        {
            hasFileDrop = e.Data.GetDataPresent(DataFormats.FileDrop);
            if (!hasFileDrop)
            {
                return null;
            }

            var paths = (string[])e.Data.GetData(DataFormats.FileDrop);
            return paths?.FirstOrDefault(IsSupportedExcelFile);
        }

        private static void ShowUnsupportedFileMessage()
        {
            MessageBox.Show(
                "엑셀 파일(.xlsx, .xls)만 드롭할 수 있습니다.",
                "알림",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private static bool IsSupportedExcelFile(string path)
        {
            var extension = Path.GetExtension(path);
            return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                   || extension.Equals(".xls", StringComparison.OrdinalIgnoreCase);
        }

        private static DataTable LoadFirstSheet(string filePath)
        {
            using var stream = File.Open(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);

            using var reader = ExcelReaderFactory.CreateReader(stream);

            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                UseColumnDataType = false,
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });

            if (dataSet.Tables.Count == 0)
            {
                return new DataTable();
            }

            return ToGridTable(dataSet.Tables[0]);
        }

        private static DataTable ToGridTable(DataTable source)
        {
            var table = new DataTable();
            var includedSourceIndexes = new List<int>();

            for (var i = 0; i < source.Columns.Count; i++)
            {
                if (!ShouldIncludeColumn(source, i))
                {
                    continue;
                }

                includedSourceIndexes.Add(i);
            }

            for (var j = 0; j < includedSourceIndexes.Count; j++)
            {
                var sourceIndex = includedSourceIndexes[j];
                var header = source.Columns[sourceIndex].ColumnName;
                if (string.IsNullOrWhiteSpace(header))
                {
                    header = $"열 {j + 1}";
                }

                var column = table.Columns.Add($"F{j}", typeof(string));
                column.ExtendedProperties[DisplayHeaderKey] = header;
                column.ExtendedProperties[MaxTextLengthKey] = header.Length;
            }

            foreach (DataRow row in source.Rows)
            {
                var newRow = table.NewRow();
                for (var j = 0; j < includedSourceIndexes.Count; j++)
                {
                    var text = FormatCellValue(row[includedSourceIndexes[j]]);
                    newRow[j] = text;

                    var column = table.Columns[j];
                    var maxLength = (int)column.ExtendedProperties[MaxTextLengthKey];
                    if (text.Length > maxLength)
                    {
                        column.ExtendedProperties[MaxTextLengthKey] = text.Length;
                    }
                }

                table.Rows.Add(newRow);
            }

            return table;
        }

        private static bool ShouldIncludeColumn(DataTable source, int columnIndex)
        {
            var header = source.Columns[columnIndex].ColumnName;
            if (!string.IsNullOrWhiteSpace(header) && !IsExcelDefaultColumnName(header))
            {
                return true;
            }

            foreach (DataRow row in source.Rows)
            {
                if (!IsCellEmpty(row[columnIndex]))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsExcelDefaultColumnName(string name)
        {
            if (!name.StartsWith("Column", StringComparison.OrdinalIgnoreCase)
                || name.Length <= "Column".Length)
            {
                return false;
            }

            return int.TryParse(name.Substring("Column".Length), out _);
        }

        private static bool IsCellEmpty(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return true;
            }

            return string.IsNullOrWhiteSpace(FormatCellValue(value));
        }

        /// <summary>
        /// 40자 초과 열은 줄바꿈 고정 너비, 이하는 SizeToCells. WrappingCellTextStyle 적용.
        /// </summary>
        private void BindDataGrid(DataTable table, DataGrid targetGrid)
        {
            targetGrid.Columns.Clear();
            targetGrid.AutoGenerateColumns = false;

            var wrappingStyle = (Style)FindResource("WrappingCellTextStyle");

            foreach (DataColumn column in table.Columns)
            {
                var header = column.ExtendedProperties[DisplayHeaderKey] as string
                             ?? column.ColumnName;
                var maxTextLength = column.ExtendedProperties[MaxTextLengthKey] is int length
                    ? length
                    : 0;
                var useWrapWidth = maxTextLength > WrapCharacterThreshold;

                var gridColumn = new DataGridTextColumn
                {
                    Header = header,
                    Binding = new Binding($"[{column.ColumnName}]")
                    {
                        Mode = BindingMode.OneWay
                    },
                    ElementStyle = wrappingStyle
                };

                if (useWrapWidth)
                {
                    gridColumn.Width = new DataGridLength(WrapColumnWidth);
                    gridColumn.MaxWidth = WrapColumnMaxWidth;
                }
                else
                {
                    gridColumn.Width = new DataGridLength(1, DataGridLengthUnitType.SizeToCells);
                    gridColumn.MinWidth = 48;
                }

                targetGrid.Columns.Add(gridColumn);
            }

            targetGrid.ItemsSource = table.DefaultView;
        }

        private static string FormatCellValue(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            switch (value)
            {
                case string s:
                    return s;
                case DateTime dt:
                    return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.CurrentCulture);
                case bool b:
                    return b ? "TRUE" : "FALSE";
                case double d:
                    return FormatNumber(d);
                case float f:
                    return FormatNumber(f);
                case decimal m:
                    return m.ToString(CultureInfo.CurrentCulture);
                default:
                    return Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty;
            }
        }

        private static string FormatNumber(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return string.Empty;
            }

            if (Math.Abs(value) < double.Epsilon)
            {
                return "0";
            }

            if (value > long.MinValue && value < long.MaxValue
                && Math.Abs(value - Math.Round(value)) < 1e-9)
            {
                return ((long)Math.Round(value)).ToString(CultureInfo.CurrentCulture);
            }

            return value.ToString("G", CultureInfo.CurrentCulture);
        }
    }
}
