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
        private const string IsResultColumnKey = "IsResultColumn";
        private const int WrapCharacterThreshold = 40;
        private const double WrapColumnWidth = 300;
        private const double MinColumnWidth = 50;
        private const double DefaultColumnWidth = 120;

        private DataGrid? _activeExcelGrid;
        private readonly Dictionary<DataGrid, int> _lastSelectedRowIndexes = new();
        private readonly Dictionary<DataGrid, string> _sourceFilePaths = new();
        private UiLogWriter? _logWriter;

        public MainWindow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
            _activeExcelGrid = LeftExcelGrid;
            _logWriter = new UiLogWriter(LogTextBox, LogScrollViewer);
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
                LogStart("BOM one row");

                var sourceTable = GetBoundTable(targetGrid);
                if (sourceTable == null || sourceTable.Rows.Count == 0)
                {
                    MessageBox.Show(
                        "변환할 데이터가 없습니다.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    LogEnd("BOM one row");
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
                    LogEnd("BOM one row");
                    return;
                }

                var totalRows = Math.Max(0, sourceTable.Rows.Count - (headerRowIndex + 3));
                LogProgress("BOM one row", $"병합 대상 행 {totalRows}건 처리 시작.");

                var mergedTable = BomOneRowTransformer.TransformWithSelectedHeaderRow(
                    sourceTable,
                    headerRowIndex,
                    (current, total) => ReportThrottledRowProgress("BOM one row", current, total, "병합 처리"));

                BindDataGrid(mergedTable, targetGrid);
                RefreshDataGridPerfect(targetGrid);
                LogProgress("BOM one row", $"결과 행 {mergedTable.Rows.Count}건 생성 완료.");
                LogEnd("BOM one row");
            }
            catch (Exception ex)
            {
                LogProgress("BOM one row", $"오류: {ex.Message}");
                LogEnd("BOM one row");
                MessageBox.Show(
                    $"BOM one row 변환 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void CompareBomMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (!TryResolveCompareGrids(out var primaryGrid, out var secondaryGrid, out var errorMessage))
            {
                MessageBox.Show(
                    errorMessage,
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                LogStart("Compare BOM");

                var primaryTable = GetBoundTable(primaryGrid);
                var secondaryTable = GetBoundTable(secondaryGrid);
                if (primaryTable == null || secondaryTable == null)
                {
                    MessageBox.Show(
                        "비교할 데이터를 읽을 수 없습니다.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    LogEnd("Compare BOM");
                    return;
                }

                var compareRowCount = Math.Max(0, primaryTable.Rows.Count - 1);
                LogProgress("Compare BOM", $"비교 대상 행 {compareRowCount}건, 참조 데이터(사양·품번 검색) 행 {Math.Max(0, secondaryTable.Rows.Count - 1)}건.");

                var comparedTable = BomCompareTransformer.CompareBom(
                    primaryTable,
                    secondaryTable,
                    (current, total) => ReportThrottledRowProgress("Compare BOM", current, total, "행 비교"));

                BindDataGrid(comparedTable, primaryGrid);
                RefreshDataGridPerfect(primaryGrid);
                LogProgress("Compare BOM", $"비교 완료. 결과 행 {comparedTable.Rows.Count}건, 복사 열: No·품번·사양·제조사·Q'ty, Result(OK/NG) 열 추가.");
                LogEnd("Compare BOM");
            }
            catch (Exception ex)
            {
                LogProgress("Compare BOM", $"오류: {ex.Message}");
                LogEnd("Compare BOM");
                MessageBox.Show(
                    $"Compare BOM 처리 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private bool TryResolveCompareGrids(
            out DataGrid primaryGrid,
            out DataGrid secondaryGrid,
            out string errorMessage)
        {
            primaryGrid = ResolveTargetGrid() ?? LeftExcelGrid;
            secondaryGrid = primaryGrid == LeftExcelGrid ? RightExcelGrid : LeftExcelGrid;
            errorMessage = string.Empty;

            if (primaryGrid.Visibility != Visibility.Visible || primaryGrid.ItemsSource == null)
            {
                errorMessage = "첫 번째(선택된) 영역에 로드된 엑셀 데이터가 없습니다. 비교할 영역을 클릭한 뒤 다시 시도해 주세요.";
                return false;
            }

            if (secondaryGrid.Visibility != Visibility.Visible || secondaryGrid.ItemsSource == null)
            {
                errorMessage = "반대편 영역에 로드된 엑셀 데이터가 없습니다. 두 영역 모두에 파일을 먼저 로드해 주세요.";
                return false;
            }

            var primaryTable = GetBoundTable(primaryGrid);
            var secondaryTable = GetBoundTable(secondaryGrid);
            if (primaryTable == null || primaryTable.Rows.Count < 2)
            {
                errorMessage = "첫 번째 데이터에 헤더 행을 제외하고 비교할 행이 없습니다.";
                return false;
            }

            if (secondaryTable == null || secondaryTable.Rows.Count < 2)
            {
                errorMessage = "두 번째 데이터에 헤더 행을 제외하고 비교할 행이 없습니다.";
                return false;
            }

            return true;
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
                LogStart("Save As");
                LogProgress("Save As", $"저장 경로: {dialog.FileName}");

                Mouse.OverrideCursor = Cursors.Wait;
                ExcelXlsxExporter.SaveDataTable(table, dialog.FileName);
                LogProgress("Save As", $"행 {table.Rows.Count}건, 열 {table.Columns.Count}건 저장 완료.");
                LogEnd("Save As");
                MessageBox.Show(
                    $"파일을 저장했습니다.\n{dialog.FileName}",
                    "저장 완료",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogProgress("Save As", $"오류: {ex.Message}");
                LogEnd("Save As");
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

        private void LogStart(string functionName)
        {
            _logWriter?.LogStart(functionName);
        }

        private void LogEnd(string functionName)
        {
            _logWriter?.LogEnd(functionName);
        }

        private void LogProgress(string functionName, string message)
        {
            _logWriter?.LogProgress(functionName, message);
        }

        private void ReportThrottledRowProgress(string functionName, int current, int total, string action)
        {
            if (!ShouldReportProgress(current, total))
            {
                return;
            }

            _logWriter?.LogRowProgress(functionName, current, total, action);
        }

        private static bool ShouldReportProgress(int current, int total)
        {
            if (total <= 0)
            {
                return false;
            }

            if (current == 1 || current == total)
            {
                return true;
            }

            var step = Math.Max(1, total / 20);
            return current % step == 0;
        }

        private bool _isBindingColumns;
        private bool _isRefreshingGrid;

        private void RefreshDataGridPerfect(DataGrid grid, bool afterHorizontalScroll = false)
        {
            if (grid.ItemsSource == null)
            {
                return;
            }

            _isRefreshingGrid = true;
            try
            {
                InvalidateDataGridViewport(grid);
                grid.UpdateLayout();
                grid.InvalidateVisual();

                if (afterHorizontalScroll)
                {
                    NudgeFirstColumnWidthSafely(grid);
                    grid.UpdateLayout();
                    grid.InvalidateVisual();
                }
            }
            finally
            {
                _isRefreshingGrid = false;
            }

            if (afterHorizontalScroll)
            {
                return;
            }

            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
            {
                if (System.Windows.Data.CollectionViewSource.GetDefaultView(grid.ItemsSource) is ICollectionView view)
                {
                    view.Refresh();
                }

                InvalidateDataGridViewport(grid);
                grid.UpdateLayout();
                grid.InvalidateVisual();
            }));
        }

        private static void InvalidateDataGridViewport(DataGrid grid)
        {
            grid.InvalidateMeasure();
            grid.InvalidateArrange();
            grid.InvalidateVisual();

            var scrollViewer = GetScrollViewer(grid);
            if (scrollViewer == null)
            {
                return;
            }

            scrollViewer.InvalidateMeasure();
            scrollViewer.InvalidateArrange();
            scrollViewer.InvalidateVisual();
        }

        private static void NudgeFirstColumnWidthSafely(DataGrid grid)
        {
            if (grid.Columns.Count == 0)
            {
                return;
            }

            var firstColumn = grid.Columns[0];
            if (!firstColumn.Width.IsAbsolute)
            {
                return;
            }

            var originalWidth = firstColumn.Width.Value;
            firstColumn.Width = new DataGridLength(originalWidth + 0.01, DataGridLengthUnitType.Pixel);
            firstColumn.Width = new DataGridLength(originalWidth, DataGridLengthUnitType.Pixel);
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
                targetGrid.ItemsSource = null;

                var table = await Task.Run(() => ExcelSheetLoader.LoadFirstSheet(filePath)).ConfigureAwait(true);

                BindDataGrid(table, targetGrid);
                _sourceFilePaths[targetGrid] = filePath;
                dropHint.Visibility = Visibility.Collapsed;
                targetGrid.Visibility = Visibility.Visible;

                await Dispatcher.InvokeAsync(
                    () => RefreshDataGridPerfect(targetGrid),
                    DispatcherPriority.ApplicationIdle);
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

        /// <summary>
        /// 40자 초과 열은 줄바꿈 고정 너비, 이하는 계산된 픽셀 너비. TemplateColumn으로 A열 렌더링 안정화.
        /// </summary>
        private void BindDataGrid(DataTable table, DataGrid targetGrid)
        {
            _isBindingColumns = true;
            try
            {
                targetGrid.ItemsSource = null;
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

                    var columnWidth = EstimateColumnWidth(maxTextLength);

                    var isResultColumn = column.ExtendedProperties.ContainsKey(IsResultColumnKey)
                        && column.ExtendedProperties[IsResultColumnKey] is true;

                    var gridColumn = new DataGridTemplateColumn
                    {
                        Header = ToExcelColumnLetter(columnIndex),
                        CellTemplate = isResultColumn
                            ? CreateResultCellTemplate(column.ColumnName)
                            : CreateCellTemplate(column.ColumnName, wrappingStyle),
                        Width = new DataGridLength(columnWidth, DataGridLengthUnitType.Pixel),
                        MinWidth = MinColumnWidth,
                        CanUserResize = true
                    };

                    if (maxTextLength > WrapCharacterThreshold)
                    {
                        gridColumn.Width = new DataGridLength(WrapColumnWidth, DataGridLengthUnitType.Pixel);
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
            if (_isBindingColumns || _isRefreshingGrid)
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
            grid.InvalidateMeasure();
        }

        private void EnforceMinColumnWidths(DataGrid grid)
        {
            foreach (var column in grid.Columns)
            {
                if (column.ActualWidth < MinColumnWidth)
                {
                    column.Width = new DataGridLength(MinColumnWidth, DataGridLengthUnitType.Pixel);
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

        private DataTemplate CreateResultCellTemplate(string columnName)
        {
            var converter = (IValueConverter)FindResource("ResultToBackgroundConverter");
            var template = new DataTemplate();
            var borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.SetValue(Border.PaddingProperty, new Thickness(8, 6, 8, 6));
            borderFactory.SetBinding(Border.BackgroundProperty, new Binding($"[{columnName}]")
            {
                Converter = converter,
                Mode = BindingMode.OneWay
            });

            var textFactory = new FrameworkElementFactory(typeof(TextBlock));
            textFactory.SetBinding(TextBlock.TextProperty, new Binding($"[{columnName}]")
            {
                Mode = BindingMode.OneWay
            });
            textFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            borderFactory.AppendChild(textFactory);
            template.VisualTree = borderFactory;
            return template;
        }

        private static double EstimateColumnWidth(int maxTextLength)
        {
            if (maxTextLength == 0)
            {
                return DefaultColumnWidth;
            }

            if (maxTextLength > WrapCharacterThreshold)
            {
                return WrapColumnWidth;
            }

            return Math.Round(Math.Clamp(maxTextLength * 7.5 + 28, MinColumnWidth, 240));
        }

        private readonly Dictionary<DataGrid, DispatcherTimer> _columnResizeTimers = new();
        private readonly Dictionary<DataGrid, DispatcherTimer> _horizontalScrollRefreshTimers = new();

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

            if (!_horizontalScrollRefreshTimers.TryGetValue(grid, out var timer))
            {
                timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(150)
                };
                timer.Tick += HorizontalScrollRefreshTimer_Tick;
                _horizontalScrollRefreshTimers[grid] = timer;
            }

            timer.Tag = grid;
            timer.Stop();
            timer.Start();
        }

        private void HorizontalScrollRefreshTimer_Tick(object? sender, EventArgs e)
        {
            if (sender is not DispatcherTimer timer)
            {
                return;
            }

            timer.Stop();

            if (timer.Tag is not DataGrid grid)
            {
                return;
            }

            RefreshDataGridPerfect(grid, afterHorizontalScroll: true);
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
    }
}
