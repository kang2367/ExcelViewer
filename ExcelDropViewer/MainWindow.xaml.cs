using System;
using System.Collections.Generic;
using System.Data;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelDataReader;
using Microsoft.Win32;

namespace ExcelDropViewer
{
    public partial class MainWindow : Window
    {
        private const string MaxTextLengthKey = "MaxTextLength";
        private const int WrapCharacterThreshold = 40;
        private const double WrapColumnWidth = 300;
        private const double WrapColumnMaxWidth = 350;
        private const double MinColumnWidth = 50;

        private DataGrid? _activeExcelGrid;
        private readonly Dictionary<DataGrid, int> _lastSelectedRowIndexes = new();
        private readonly Dictionary<DataGrid, string> _sourceFilePaths = new();

        public MainWindow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
            _activeExcelGrid = LeftExcelGrid;
        }

        private void ExcelGrid_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                _activeExcelGrid = grid;
            }
        }

        private void BomOneRowMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var targetGrid = ResolveTargetGrid();
            if (targetGrid == null)
            {
                MessageBox.Show(
                    "변환할 DataGrid를 선택할 수 없습니다. 좌측 또는 우측 영역을 클릭한 뒤 다시 시도해 주세요.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            if (targetGrid.Visibility != Visibility.Visible || targetGrid.ItemsSource == null)
            {
                MessageBox.Show(
                    "선택한 영역에 로드된 엑셀 데이터가 없습니다.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                var sourceTable = GetBoundTable(targetGrid);
                if (sourceTable == null || sourceTable.Rows.Count == 0)
                {
                    MessageBox.Show(
                        "변환할 데이터가 없습니다.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                var headerRowIndex = ResolveHeaderRowIndex(targetGrid);
                if (headerRowIndex < 0)
                {
                    MessageBox.Show(
                        "헤더로 지정할 행을 먼저 선택해 주세요.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                var mergedTable = BomOneRowTransformer.TransformWithSelectedHeaderRow(sourceTable, headerRowIndex);
                BindDataGrid(mergedTable, targetGrid);
                RefreshGridAfterDataTransform(targetGrid);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"BOM one row 변환 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SaveAsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var targetGrid = ResolveTargetGrid();
            if (targetGrid == null
                || targetGrid.Visibility != Visibility.Visible
                || targetGrid.ItemsSource == null)
            {
                MessageBox.Show(
                    "저장할 데이터가 없습니다. 좌측 또는 우측 영역에 엑셀 파일을 먼저 로드해 주세요.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var table = GetBoundTable(targetGrid);
            if (table == null || table.Rows.Count == 0)
            {
                MessageBox.Show(
                    "저장할 데이터가 없습니다.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            _sourceFilePaths.TryGetValue(targetGrid, out var sourceFilePath);

            var dialog = new SaveFileDialog
            {
                Title = "다른 이름으로 저장",
                Filter = "Excel 파일 (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                FileName = BuildDefaultSaveFileName(sourceFilePath)
            };

            var sourceDirectory = GetSourceDirectory(sourceFilePath);
            if (!string.IsNullOrWhiteSpace(sourceDirectory))
            {
                dialog.InitialDirectory = sourceDirectory;
            }

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                ExcelXlsxExporter.SaveDataTable(table, dialog.FileName);
                MessageBox.Show(
                    $"파일을 저장했습니다.\n{dialog.FileName}",
                    "저장 완료",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"파일 저장 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private DataGrid? ResolveTargetGrid()
        {
            if (_activeExcelGrid is { Visibility: Visibility.Visible, ItemsSource: not null })
            {
                return _activeExcelGrid;
            }

            if (LeftExcelGrid.Visibility == Visibility.Visible && LeftExcelGrid.ItemsSource != null)
            {
                return LeftExcelGrid;
            }

            if (RightExcelGrid.Visibility == Visibility.Visible && RightExcelGrid.ItemsSource != null)
            {
                return RightExcelGrid;
            }

            return _activeExcelGrid;
        }

        private static DataTable? GetBoundTable(DataGrid grid)
        {
            return grid.ItemsSource switch
            {
                DataView dataView => dataView.Table,
                DataTable dataTable => dataTable,
                _ => null
            };
        }

        private void ExcelGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not DataGrid grid)
            {
                return;
            }

            var rowIndex = TryGetSelectedRowIndex(grid, out _);
            if (rowIndex >= 0)
            {
                _lastSelectedRowIndexes[grid] = rowIndex;
            }
        }

        private int ResolveHeaderRowIndex(DataGrid grid)
        {
            var rowIndex = TryGetSelectedRowIndex(grid, out _);
            if (rowIndex >= 0)
            {
                return rowIndex;
            }

            if (_lastSelectedRowIndexes.TryGetValue(grid, out rowIndex) && rowIndex >= 0)
            {
                return rowIndex;
            }

            return -1;
        }

        private static int TryGetSelectedRowIndex(DataGrid grid, out DataRowView? rowView)
        {
            rowView = null;

            if (grid.ItemsSource is not DataView)
            {
                return -1;
            }

            if (grid.CurrentItem is DataRowView currentRow)
            {
                rowView = currentRow;
            }
            else if (grid.SelectedItem is DataRowView selectedRow)
            {
                rowView = selectedRow;
            }
            else if (grid.SelectedCells.Count > 0 && grid.SelectedCells[0].Item is DataRowView cellRow)
            {
                rowView = cellRow;
            }

            if (rowView == null)
            {
                return -1;
            }

            return rowView.Row.Table.Rows.IndexOf(rowView.Row);
        }

        private void RefreshGridAfterDataTransform(DataGrid grid)
        {
            grid.RowHeight = double.NaN;
            grid.UpdateLayout();

            var scrollViewer = GetScrollViewer(grid);
            if (scrollViewer != null)
            {
                RefreshGridLayout(grid, scrollViewer, fullRepair: true);
            }
            else
            {
                grid.InvalidateMeasure();
                grid.InvalidateArrange();
                grid.UpdateLayout();
            }

            Dispatcher.BeginInvoke(() =>
            {
                if (System.Windows.Data.CollectionViewSource.GetDefaultView(grid.ItemsSource) is ICollectionView view)
                {
                    view.Refresh();
                }

                grid.InvalidateMeasure();
                grid.UpdateLayout();
            }, DispatcherPriority.Loaded);
        }

        private void ExcelGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString(CultureInfo.InvariantCulture);
        }

        private void ExcelGrid_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.Shift || sender is not DataGrid grid)
            {
                return;
            }

            var scrollViewer = GetScrollViewer(grid);
            if (scrollViewer == null)
            {
                return;
            }

            var nextOffset = scrollViewer.HorizontalOffset - e.Delta;
            scrollViewer.ScrollToHorizontalOffset(Math.Max(0, nextOffset));
            e.Handled = true;
        }

        private static ScrollViewer? GetScrollViewer(DependencyObject element)
        {
            if (element is ScrollViewer scrollViewer)
            {
                return scrollViewer;
            }

            for (var i = 0; i < VisualTreeHelper.GetChildrenCount(element); i++)
            {
                var child = VisualTreeHelper.GetChild(element, i);
                var result = GetScrollViewer(child);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
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
                _sourceFilePaths[targetGrid] = filePath;
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

        private static string BuildDefaultSaveFileName(string? sourceFilePath)
        {
            var baseName = string.IsNullOrWhiteSpace(sourceFilePath)
                ? "export"
                : Path.GetFileNameWithoutExtension(sourceFilePath);
            var dateText = DateTime.Today.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            return $"{baseName}_Modify_{dateText}.xlsx";
        }

        private static string? GetSourceDirectory(string? sourceFilePath)
        {
            if (string.IsNullOrWhiteSpace(sourceFilePath))
            {
                return null;
            }

            var directory = Path.GetDirectoryName(sourceFilePath);
            return string.IsNullOrWhiteSpace(directory) ? null : directory;
        }

        private static string? TryGetDroppedExcelPath(DragEventArgs e, out bool hasFileDrop)
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
                    UseHeaderRow = false
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
                var column = table.Columns.Add($"F{j}", typeof(string));
                column.ExtendedProperties[MaxTextLengthKey] = 0;
            }

            foreach (DataRow row in source.Rows)
            {
                var newRow = table.NewRow();
                for (var j = 0; j < includedSourceIndexes.Count; j++)
                {
                    var text = FormatCellValue(row[includedSourceIndexes[j]]);
                    newRow[j] = text;

                    var column = table.Columns[j];
                    var maxLength = column.ExtendedProperties[MaxTextLengthKey] is int currentMax
                        ? currentMax
                        : 0;
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
        /// 40자 초과 열은 줄바꿈 고정 너비, 이하는 계산된 픽셀 너비. TemplateColumn으로 A열 렌더링 안정화.
        /// </summary>
        private bool _isBindingColumns;

        private void BindDataGrid(DataTable table, DataGrid targetGrid)
        {
            _isBindingColumns = true;
            try
            {
                targetGrid.Columns.Clear();
                targetGrid.AutoGenerateColumns = false;
                targetGrid.FrozenColumnCount = 0;

                var wrappingStyle = (Style)FindResource("WrappingCellTextStyle");

                for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    var column = table.Columns[columnIndex];
                    var maxTextLength = column.ExtendedProperties[MaxTextLengthKey] is int length
                        ? length
                        : 0;

                    var gridColumn = new DataGridTemplateColumn
                    {
                        Header = ToExcelColumnLetter(columnIndex),
                        CellTemplate = CreateCellTemplate(column.ColumnName, wrappingStyle),
                        Width = new DataGridLength(EstimateColumnWidth(maxTextLength)),
                        MinWidth = MinColumnWidth,
                        CanUserResize = true
                    };

                    if (maxTextLength > WrapCharacterThreshold)
                    {
                        gridColumn.MaxWidth = WrapColumnMaxWidth;
                    }

                    targetGrid.Columns.Add(gridColumn);
                    WireColumnResize(targetGrid, gridColumn);
                }

                targetGrid.ItemsSource = table.DefaultView;
            }
            finally
            {
                _isBindingColumns = false;
            }
        }

        private void WireColumnResize(DataGrid grid, DataGridColumn column)
        {
            var descriptor = DependencyPropertyDescriptor.FromProperty(
                DataGridColumn.WidthProperty,
                typeof(DataGridColumn));

            descriptor?.AddValueChanged(column, (_, _) => ScheduleColumnResizeRefresh(grid));
        }

        /// <summary>
        /// 드래그 중에는 타이머만 재시작하고, 무거운 갱신은 하지 않습니다.
        /// </summary>
        private void ScheduleColumnResizeRefresh(DataGrid grid)
        {
            if (_isBindingColumns)
            {
                return;
            }

            if (!_columnResizeTimers.TryGetValue(grid, out var timer))
            {
                timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(200)
                };
                timer.Tick += ColumnResizeTimer_Tick;
                _columnResizeTimers[grid] = timer;
            }

            timer.Tag = grid;
            timer.Stop();
            timer.Start();
        }

        private void ColumnResizeTimer_Tick(object? sender, EventArgs e)
        {
            if (sender is not DispatcherTimer timer || timer.Tag is not DataGrid grid)
            {
                return;
            }

            timer.Stop();
            CompleteColumnResize(grid);
        }

        private void ExcelDataGrid_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGrid grid)
            {
                return;
            }

            if (_columnResizeTimers.TryGetValue(grid, out var timer) && timer.IsEnabled)
            {
                timer.Stop();
                CompleteColumnResize(grid);
            }
        }

        private void CompleteColumnResize(DataGrid grid)
        {
            EnforceMinColumnWidths(grid);

            var scrollViewer = GetScrollViewer(grid);
            if (scrollViewer != null)
            {
                RefreshGridLayout(grid, scrollViewer, fullRepair: false);
            }
        }

        private void EnforceMinColumnWidths(DataGrid grid)
        {
            foreach (var column in grid.Columns)
            {
                if (column.ActualWidth < MinColumnWidth)
                {
                    column.Width = new DataGridLength(MinColumnWidth);
                }
            }
        }

        private static DataTemplate CreateCellTemplate(string columnName, Style textStyle)
        {
            var template = new DataTemplate();
            var factory = new FrameworkElementFactory(typeof(TextBlock));
            factory.SetBinding(TextBlock.TextProperty, new Binding($"[{columnName}]")
            {
                Mode = BindingMode.OneWay
            });
            factory.SetValue(TextBlock.StyleProperty, textStyle);
            factory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
            factory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Top);
            template.VisualTree = factory;
            return template;
        }

        private static double EstimateColumnWidth(int maxTextLength)
        {
            if (maxTextLength > WrapCharacterThreshold)
            {
                return WrapColumnWidth;
            }

            return Math.Clamp(maxTextLength * 7.5 + 28, MinColumnWidth, 240);
        }

        private readonly Dictionary<DataGrid, DispatcherTimer> _horizontalScrollTimers = new();
        private readonly Dictionary<DataGrid, DispatcherTimer> _columnResizeTimers = new();

        private void ExcelDataGrid_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (Math.Abs(e.HorizontalChange) < 0.01)
            {
                return;
            }

            ScrollViewer? scrollViewer = sender as ScrollViewer;
            if (scrollViewer == null && sender is DataGrid dataGrid)
            {
                scrollViewer = GetScrollViewer(dataGrid);
            }

            if (scrollViewer == null)
            {
                return;
            }

            var grid = sender as DataGrid ?? FindParent<DataGrid>(scrollViewer);
            if (grid == null)
            {
                return;
            }

            if (!_horizontalScrollTimers.TryGetValue(grid, out var timer))
            {
                timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(120)
                };
                timer.Tick += HorizontalScrollTimer_Tick;
                _horizontalScrollTimers[grid] = timer;
            }

            timer.Tag = scrollViewer;
            timer.Stop();
            timer.Start();
        }

        private void HorizontalScrollTimer_Tick(object? sender, EventArgs e)
        {
            if (sender is not DispatcherTimer timer)
            {
                return;
            }

            timer.Stop();

            if (timer.Tag is not ScrollViewer scrollViewer)
            {
                return;
            }

            var grid = FindParent<DataGrid>(scrollViewer);
            if (grid == null)
            {
                return;
            }

            RefreshGridLayout(grid, scrollViewer, fullRepair: true);
        }

        /// <param name="fullRepair">true: 가로 스크롤 후 셀 깨짐 복구용 전체 갱신. false: 열 리사이즈 완료 후 최소 레이아웃 갱신.</param>
        private static void RefreshGridLayout(DataGrid grid, ScrollViewer scrollViewer, bool fullRepair)
        {
            var horizontalOffset = scrollViewer.HorizontalOffset;
            var verticalOffset = scrollViewer.VerticalOffset;

            if (fullRepair)
            {
                Keyboard.ClearFocus();
                grid.UnselectAll();

                for (var i = 0; i < grid.Items.Count; i++)
                {
                    if (grid.ItemContainerGenerator.ContainerFromIndex(i) is not DataGridRow row)
                    {
                        continue;
                    }

                    row.Header = (i + 1).ToString(CultureInfo.InvariantCulture);
                    row.InvalidateArrange();
                    row.InvalidateVisual();
                }

                if (System.Windows.Data.CollectionViewSource.GetDefaultView(grid.ItemsSource) is ICollectionView view)
                {
                    view.Refresh();
                }
            }

            grid.InvalidateMeasure();
            grid.InvalidateArrange();
            grid.UpdateLayout();

            scrollViewer.ScrollToHorizontalOffset(horizontalOffset);
            scrollViewer.ScrollToVerticalOffset(verticalOffset);
        }

        private static T? FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            var current = VisualTreeHelper.GetParent(child);
            while (current != null)
            {
                if (current is T match)
                {
                    return match;
                }

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        private static string ToExcelColumnLetter(int columnIndex)
        {
            var dividend = columnIndex + 1;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
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
