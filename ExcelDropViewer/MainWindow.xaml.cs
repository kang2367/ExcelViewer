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
using System.Windows.Input;
using System.Windows.Forms.Integration;
using Microsoft.Win32;
using unvell.ReoGrid;
using unvell.ReoGrid.IO;

namespace ExcelDropViewer
{
    public partial class MainWindow : Window
    {
        private readonly ReoGridControl LeftReoGrid;
        private readonly ReoGridControl RightReoGrid;
        private ReoGridControl? _activeReoGrid;
        private readonly Dictionary<ReoGridControl, int> _lastSelectedRowIndexes = new();
        private readonly Dictionary<ReoGridControl, string> _sourceFilePaths = new();
        private readonly HashSet<ReoGridControl> _loadedReoGrids = new();
        private UiLogWriter? _logWriter;
        private readonly DigiKeyApiService _digiKeyApiService = new();
        private string? _contextMenuPartNumber;
        private string? _contextMenuManufacturer;

        public MainWindow()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            InitializeComponent();
            LeftReoGrid = CreateReoGrid(LeftReoGridHost);
            RightReoGrid = CreateReoGrid(RightReoGridHost);
            _activeReoGrid = LeftReoGrid;
            _logWriter = new UiLogWriter(LogTextBox, LogScrollViewer);
        }

        private ReoGridControl CreateReoGrid(WindowsFormsHost host)
        {
            var grid = new ReoGridControl
            {
                Readonly = true,
                ShowScrollEndSpacing = true,
                Dock = System.Windows.Forms.DockStyle.Fill
            };
            grid.GotFocus += (_, _) =>
            {
                _activeReoGrid = grid;
                TrackSelectedRow(grid);
            };
            SetupReoGridContextMenu(grid);
            host.Child = grid;
            return grid;
        }

        private void SetupReoGridContextMenu(ReoGridControl grid)
        {
            var partSearchMenu = new System.Windows.Forms.ToolStripMenuItem("부품 검색");
            var digiKeyMenu = new System.Windows.Forms.ToolStripMenuItem("Digi-Key");
            digiKeyMenu.Click += async (_, _) => await MenuDigiKeySearchAsync();
            partSearchMenu.DropDownItems.Add(digiKeyMenu);

            var contextMenu = new System.Windows.Forms.ContextMenuStrip();
            contextMenu.Items.Add(partSearchMenu);
            contextMenu.Opening += (_, _) => CaptureContextMenuPartNumber(grid);
            grid.ContextMenuStrip = contextMenu;
        }

        private void CaptureContextMenuPartNumber(ReoGridControl grid)
        {
            _activeReoGrid = grid;
            var sheet = grid.CurrentWorksheet;
            var range = sheet.SelectionRange;
            if (range.Row < 0 || range.Col < 0)
            {
                _contextMenuPartNumber = null;
                _contextMenuManufacturer = null;
                return;
            }

            _contextMenuPartNumber = ReoGridWorksheetAdapter.GetCellText(sheet, range.Row, range.Col);
            _contextMenuManufacturer = ReoGridWorksheetAdapter.GetCellText(sheet, range.Row, range.Col + 1);
        }

        private async Task MenuDigiKeySearchAsync()
        {
            var partNumber = _contextMenuPartNumber?.Trim();
            if (string.IsNullOrWhiteSpace(partNumber))
            {
                LogProgress("Digi-Key", "검색할 부품 번호가 선택되지 않았습니다.");
                return;
            }

            var manufacturer = DigiKeyManufacturerNameNormalizer.Normalize(_contextMenuManufacturer);

            var config = DigikeyConfigStore.Load();
            if (string.IsNullOrWhiteSpace(config.ClientId) || string.IsNullOrWhiteSpace(config.ClientSecret))
            {
                LogProgress(
                    "Digi-Key",
                    "CONFIG.INI에 Digi-Key API 설정이 없습니다. [설정] > [Digikey] 메뉴에서 먼저 설정해 주세요.");
                return;
            }

            try
            {
                LogStart("Digi-Key");
                LogProgress(
                    "Digi-Key",
                    string.IsNullOrWhiteSpace(manufacturer)
                        ? $"부품 번호 '{partNumber}' 검색 요청 중..."
                        : $"부품 번호 '{partNumber}', 제조사 '{manufacturer}' 검색 요청 중...");

                Mouse.OverrideCursor = Cursors.Wait;
                var summary = await _digiKeyApiService.SearchByPartNumberAsync(partNumber, manufacturer, config);
                _logWriter?.LogMultiline(DigiKeySearchResultFormatter.Format(summary));
                UpsertDigiKeySearchResultToRightGrid(summary);
                LogProgress("Digi-Key", "우측 그리드에 검색 결과를 반영했습니다.");
                LogEnd("Digi-Key");
            }
            catch (Exception ex)
            {
                LogProgress("Digi-Key", $"오류: {ex.Message}");
                LogEnd("Digi-Key");
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void UpsertDigiKeySearchResultToRightGrid(DigiKeyProductSummary summary)
        {
            EnsureRightGridVisible();
            DigiKeySearchResultGridWriter.Upsert(RightReoGrid.CurrentWorksheet, summary);
            _loadedReoGrids.Add(RightReoGrid);
        }

        private void EnsureRightGridVisible()
        {
            RightDropHint.Visibility = Visibility.Collapsed;
            RightReoGridHost.Visibility = Visibility.Visible;
        }

        private void TrackSelectedRow(ReoGridControl grid)
        {
            var rowIndex = ReoGridWorksheetAdapter.TryGetSelectedRowIndex(grid);
            if (rowIndex >= 0)
            {
                _lastSelectedRowIndexes[grid] = rowIndex;
            }
        }

        private void AttachWorksheetSelectionTracking(ReoGridControl grid)
        {
            var sheet = grid.CurrentWorksheet;
            sheet.SelectionRangeChanged += (_, _) => TrackSelectedRow(grid);
        }

        private void BomOneRowMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var targetGrid = ResolveTargetGrid();
            if (!TryGetLoadedWorksheet(targetGrid, out var worksheet, out var validationMessage))
            {
                MessageBox.Show(
                    validationMessage,
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                LogStart("BOM one row");

                var sourceTable = ReoGridWorksheetAdapter.ToDataTable(worksheet);
                if (sourceTable.Rows.Count == 0)
                {
                    MessageBox.Show(
                        "변환할 데이터가 없습니다.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    LogEnd("BOM one row");
                    return;
                }

                var headerRowIndex = ResolveHeaderRowIndex(targetGrid!);
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

                ReoGridWorksheetAdapter.ApplyDataTable(worksheet, mergedTable);
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

        private void MakeBomDbMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var targetGrid = ResolveTargetGrid();
            if (!TryGetLoadedWorksheet(targetGrid, out var worksheet, out var validationMessage))
            {
                MessageBox.Show(
                    validationMessage,
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                LogStart("Make BOM DB");

                var sourceTable = ReoGridWorksheetAdapter.ToDataTable(worksheet);
                if (sourceTable.Rows.Count == 0)
                {
                    MessageBox.Show(
                        "저장할 데이터가 없습니다.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    LogEnd("Make BOM DB");
                    return;
                }

                var headerRowIndex = ResolveHeaderRowIndex(targetGrid!);
                var dataRowCount = headerRowIndex >= 0
                    ? Math.Max(0, sourceTable.Rows.Count - (headerRowIndex + 1))
                    : Math.Max(0, sourceTable.Rows.Count - 1);

                LogProgress("Make BOM DB", $"DB 저장 경로: {BomDbPaths.GetDatabasePath()}");
                LogProgress("Make BOM DB", $"BOM DB 적재 대상 행 최대 {dataRowCount}건.");

                Mouse.OverrideCursor = Cursors.Wait;
                var result = BomDbImporter.Import(
                    sourceTable,
                    headerRowIndex,
                    this,
                    (current, total) => ReportThrottledRowProgress("Make BOM DB", current, total, "DB 적재"));

                LogProgress("Make BOM DB", $"헤더 행: {result.HeaderRowIndex + 1}행, DB 경로: {result.DatabasePath}");
                if (!string.IsNullOrWhiteSpace(result.BackupFileName))
                {
                    LogProgress("Make BOM DB", $"백업 파일: {result.BackupFileName}");
                }

                LogProgress(
                    "Make BOM DB",
                    $"신규 추가: {result.InsertedCount}건, 업데이트: {result.UpdatedCount}건, 건너뜀: {result.SkippedCount}건{(result.Cancelled ? ", 취소됨(롤백)" : string.Empty)}.");
                LogEnd("Make BOM DB");

                MessageBox.Show(
                    result.BuildSummaryMessage(),
                    "Make BOM DB",
                    MessageBoxButton.OK,
                    result.Cancelled ? MessageBoxImage.Warning : MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogProgress("Make BOM DB", $"오류: {ex.Message}");
                LogEnd("Make BOM DB");
                MessageBox.Show(
                    $"Make BOM DB 처리 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void ReferenceCheckMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var targetGrid = ResolveTargetGrid();
            if (!TryGetLoadedWorksheet(targetGrid, out var worksheet, out var validationMessage))
            {
                MessageBox.Show(
                    validationMessage,
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                LogStart("Reference check");

                var sourceTable = ReoGridWorksheetAdapter.ToDataTable(worksheet);
                if (sourceTable.Rows.Count < 2)
                {
                    MessageBox.Show(
                        "헤더 행을 제외하고 확인할 데이터 행이 없습니다.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    LogEnd("Reference check");
                    return;
                }

                var dataRowCount = Math.Max(0, sourceTable.Rows.Count - 1);
                LogProgress("Reference check", $"Reference 열 연속성 확인 대상 행 {dataRowCount}건.");

                var report = ReferenceCheckTransformer.CheckContinuity(
                    sourceTable,
                    (current, total) => ReportThrottledRowProgress("Reference check", current, total, "행 확인"));

                foreach (var line in report.ReportLines)
                {
                    LogProgress("Reference check", line);
                }

                LogEnd("Reference check");
                ScrollableMessageDialog.Show(this, "Reference check", report.MessageText);
            }
            catch (Exception ex)
            {
                LogProgress("Reference check", $"오류: {ex.Message}");
                LogEnd("Reference check");
                MessageBox.Show(
                    $"Reference check 처리 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void MakePbaBomMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (!TryGetLoadedWorksheet(LeftReoGrid, out var rawWorksheet, out _)
                || !TryGetLoadedWorksheet(RightReoGrid, out var templateWorksheet, out _))
            {
                MessageBox.Show(
                    "왼쪽 창의 Raw Data 파일과 오른쪽 창의 Template 파일을 모두 열어주세요.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                LogStart("Make PBA BOM");

                var rawTable = ReoGridWorksheetAdapter.ToDataTable(rawWorksheet);
                var templateTable = ReoGridWorksheetAdapter.ToDataTable(templateWorksheet);

                if (rawTable.Rows.Count == 0 || templateTable.Rows.Count == 0)
                {
                    MessageBox.Show(
                        "왼쪽 창의 Raw Data 파일과 오른쪽 창의 Template 파일을 모두 열어주세요.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    LogEnd("Make PBA BOM");
                    return;
                }

                LogProgress("Make PBA BOM", "Raw Data를 Template BOM 양식으로 변환 중...");

                var result = PbaBomTransformer.Transform(
                    rawTable,
                    templateTable,
                    (current, total) => ReportThrottledRowProgress("Make PBA BOM", current, total, "변환 처리"));

                var sampleColumn = result.Layout.ReferenceColumn >= 0
                    ? result.Layout.ReferenceColumn
                    : 0;
                var baseFontSize = PbaBomWorksheetStyler.TryGetBaseFontSize(
                    templateWorksheet,
                    result.Layout.TemplateHeaderRow + 1,
                    sampleColumn);

                ReoGridWorksheetAdapter.ApplyDataTable(templateWorksheet, result.OutputTable);
                PbaBomWorksheetStyler.Apply(templateWorksheet, result.Layout, baseFontSize);
                EnsureRightGridVisible();
                _loadedReoGrids.Add(RightReoGrid);

                LogProgress(
                    "Make PBA BOM",
                    $"PBA BOM 변환이 성공적으로 완료되었습니다. (총 {result.NormalItemCount}개 항목, NC {result.NcItemCount}개 항목)");
                LogEnd("Make PBA BOM");
            }
            catch (Exception ex)
            {
                LogProgress("Make PBA BOM", $"오류: {ex.Message}");
                LogEnd("Make PBA BOM");
                MessageBox.Show(
                    $"Make PBA BOM 처리 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void MatchPdbBomWithDbMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var targetGrid = ResolveTargetGrid();
            if (!TryGetLoadedWorksheet(targetGrid, out var worksheet, out _))
            {
                MessageBox.Show(
                    "매칭 작업을 진행할 BOM 파일이 열려있지 않습니다.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var databasePath = BomDbPaths.GetDatabasePath();
            if (!File.Exists(databasePath))
            {
                MessageBox.Show(
                    "Data/BOM_Master.db 파일을 찾을 수 없습니다.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                LogStart("Match PDB BOM with DB");

                var sourceTable = ReoGridWorksheetAdapter.ToDataTable(worksheet);
                if (sourceTable.Rows.Count == 0)
                {
                    MessageBox.Show(
                        "매칭 작업을 진행할 BOM 파일이 열려있지 않습니다.",
                        "알림",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    LogEnd("Match PDB BOM with DB");
                    return;
                }

                LogProgress("Match PDB BOM with DB", $"DB 경로: {databasePath}");
                Mouse.OverrideCursor = Cursors.Wait;

                var result = PdbBomDbMatcher.Match(
                    sourceTable,
                    databasePath,
                    (current, total) => ReportThrottledRowProgress(
                        "Match PDB BOM with DB",
                        current,
                        total,
                        "DB 매칭"));

                ReoGridWorksheetAdapter.ApplyDataTable(worksheet, result.OutputTable);
                PdbBomDbMatchWorksheetStyler.Apply(worksheet, result.Layout);

                LogProgress(
                    "Match PDB BOM with DB",
                    $"PDB BOM DB 매칭 완료: 총 {result.TotalProcessed}개 항목 중 {result.MatchedCount}개 매칭 성공, {result.UnmatchedCount}개 미매칭(DB 미등록)");
                LogEnd("Match PDB BOM with DB");
            }
            catch (Exception ex)
            {
                LogProgress("Match PDB BOM with DB", $"오류: {ex.Message}");
                LogEnd("Match PDB BOM with DB");
                MessageBox.Show(
                    $"Match PDB BOM with DB 처리 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
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

                var primaryWorksheet = primaryGrid.CurrentWorksheet;
                var secondaryWorksheet = secondaryGrid.CurrentWorksheet;
                var primaryTable = ReoGridWorksheetAdapter.ToDataTable(primaryWorksheet);
                var secondaryTable = ReoGridWorksheetAdapter.ToDataTable(secondaryWorksheet);

                if (primaryTable.Rows.Count < 2 || secondaryTable.Rows.Count < 2)
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

                ReoGridWorksheetAdapter.ApplyDataTable(primaryWorksheet, comparedTable);
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
            out ReoGridControl primaryGrid,
            out ReoGridControl secondaryGrid,
            out string errorMessage)
        {
            primaryGrid = ResolveTargetGrid() ?? LeftReoGrid;
            secondaryGrid = primaryGrid == LeftReoGrid ? RightReoGrid : LeftReoGrid;
            errorMessage = string.Empty;

            if (!TryGetLoadedWorksheet(primaryGrid, out var primaryWorksheet, out errorMessage))
            {
                errorMessage = "첫 번째(선택된) 영역에 로드된 엑셀 데이터가 없습니다. 비교할 영역을 클릭한 뒤 다시 시도해 주세요.";
                return false;
            }

            if (!TryGetLoadedWorksheet(secondaryGrid, out var secondaryWorksheet, out _))
            {
                errorMessage = "반대편 영역에 로드된 엑셀 데이터가 없습니다. 두 영역 모두에 파일을 먼저 로드해 주세요.";
                return false;
            }

            var primaryTable = ReoGridWorksheetAdapter.ToDataTable(primaryWorksheet);
            var secondaryTable = ReoGridWorksheetAdapter.ToDataTable(secondaryWorksheet);
            if (primaryTable.Rows.Count < 2)
            {
                errorMessage = "첫 번째 데이터에 헤더 행을 제외하고 비교할 행이 없습니다.";
                return false;
            }

            if (secondaryTable.Rows.Count < 2)
            {
                errorMessage = "두 번째 데이터에 헤더 행을 제외하고 비교할 행이 없습니다.";
                return false;
            }

            return true;
        }

        private void MenuDigikeyConfig_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new DigikeyConfigWindow
            {
                Owner = this
            };
            dlg.ShowDialog();
        }

        private void SaveAsMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var targetGrid = ResolveTargetGrid();
            if (!TryGetLoadedWorksheet(targetGrid, out _, out var validationMessage))
            {
                MessageBox.Show(
                    validationMessage,
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            _sourceFilePaths.TryGetValue(targetGrid!, out var sourceFilePath);

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
                targetGrid.Save(dialog.FileName, FileFormat.Excel2007);
                LogProgress("Save As", "ReoGrid 워크시트 저장 완료.");
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

        private ReoGridControl? ResolveTargetGrid()
        {
            if (_activeReoGrid != null && _loadedReoGrids.Contains(_activeReoGrid))
            {
                return _activeReoGrid;
            }

            if (_loadedReoGrids.Contains(LeftReoGrid))
            {
                return LeftReoGrid;
            }

            if (_loadedReoGrids.Contains(RightReoGrid))
            {
                return RightReoGrid;
            }

            return _activeReoGrid;
        }

        private bool TryGetLoadedWorksheet(
            ReoGridControl? grid,
            out Worksheet worksheet,
            out string message)
        {
            worksheet = null!;
            message = string.Empty;

            if (grid == null)
            {
                message = "작업할 영역을 선택할 수 없습니다. 좌측 또는 우측 영역을 클릭한 뒤 다시 시도해 주세요.";
                return false;
            }

            if (!_loadedReoGrids.Contains(grid))
            {
                message = "선택한 영역에 로드된 엑셀 데이터가 없습니다.";
                return false;
            }

            worksheet = grid.CurrentWorksheet;
            if (!ReoGridWorksheetAdapter.HasWorksheetData(worksheet))
            {
                message = "선택한 영역에 로드된 엑셀 데이터가 없습니다.";
                return false;
            }

            return true;
        }

        private int ResolveHeaderRowIndex(ReoGridControl grid)
        {
            var rowIndex = ReoGridWorksheetAdapter.TryGetSelectedRowIndex(grid);
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

        private void LeftZone_DragOver(object sender, System.Windows.DragEventArgs e)
        {
            HandleZoneDragOver(e);
        }

        private void RightZone_DragOver(object sender, System.Windows.DragEventArgs e)
        {
            HandleZoneDragOver(e);
        }

        private async void LeftZone_Drop(object sender, System.Windows.DragEventArgs e)
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

            await LoadExcelAsync(excelPath, LeftReoGrid, LeftDropHint);
        }

        private async void RightZone_Drop(object sender, System.Windows.DragEventArgs e)
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

            await LoadExcelAsync(excelPath, RightReoGrid, RightDropHint);
        }

        private static void HandleZoneDragOver(System.Windows.DragEventArgs e)
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

        private async Task LoadExcelAsync(string filePath, ReoGridControl targetGrid, TextBlock dropHint)
        {
            var host = targetGrid == LeftReoGrid ? LeftReoGridHost : RightReoGridHost;

            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                await Task.Yield();

                targetGrid.Reset();
                targetGrid.Load(filePath);
                AttachWorksheetSelectionTracking(targetGrid);

                _sourceFilePaths[targetGrid] = filePath;
                _loadedReoGrids.Add(targetGrid);
                dropHint.Visibility = Visibility.Collapsed;
                host.Visibility = Visibility.Visible;
                _activeReoGrid = targetGrid;
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

        private static string? TryGetDroppedExcelPath(System.Windows.DragEventArgs e, out bool hasFileDrop)
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
    }
}
