# ExcelDropViewer

WPF 앱으로 엑셀 파일(.xlsx, .xls)을 좌·우 영역에 드래그 앤 드롭하면 첫 번째 시트 데이터를 DataGrid에 표시합니다.

**요구 사항:** .NET 10 SDK

## 실행

```powershell
cd C:\Users\USER\ExcelDropViewer\ExcelDropViewer
dotnet run
```

## 기능

- 전체 화면 연한 회색 배경 (`#F5F6F8`)
- 엑셀 파일 드래그 시 **복사(Copy)** 커서
- 드롭 후 읽기 전용 DataGrid (행 가상화 활성화)
