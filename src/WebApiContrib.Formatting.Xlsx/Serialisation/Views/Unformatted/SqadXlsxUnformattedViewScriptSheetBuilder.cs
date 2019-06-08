using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System.Data;
using System.Linq;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewScriptSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly string _dataUrl;
        private readonly bool _needCreatePivotSheet;

        public SqadXlsxUnformattedViewScriptSheetBuilder(string dataUrl, bool needCreatePivotSheet)
            :base(ExportViewConstants.UnformattedViewScriptSheetName, shouldAutoFit: false)
        {
            _dataUrl = dataUrl;
            _needCreatePivotSheet = needCreatePivotSheet;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }

            var dataScript = string.Empty;
            var refreshDataScript = string.Empty;
            var dataSheet = worksheet.Workbook
                                     .Worksheets
                                     .First(x => x.Name == ExportViewConstants.UnformattedViewDataSheetName);

            if (_dataUrl != null)
            {
                dataScript = GetDataScript(dataSheet);
                refreshDataScript = GetRefreshDataScript();
            }

            var pivotScript = string.Empty;
            if (_needCreatePivotSheet)
            {
                pivotScript = GetPivotScript(dataSheet.Dimension.ToString());
            }

            var code = $@"
Private Sub Workbook_Open()
    Dim tmpSheet As Worksheet
    Set tmpSheet = Sheets(""{ExportViewConstants.UnformattedViewScriptSheetName}"")
    If tmpSheet.Visible = xlSheetVeryHidden Then
        Exit Sub
    End If
    
    {dataScript}
    {pivotScript}

    tmpSheet.Visible = xlSheetVeryHidden
End Sub

{refreshDataScript}
";
           worksheet.Workbook.CodeModule.Code = code;
        }

        private string GetDataScript(ExcelWorksheet worksheet)
        {
            const string cellsName = "nwshp_hl_en_tab_wn";
            var cells = worksheet.Cells[worksheet.Dimension.Address];
            worksheet.Names.Add(cellsName, cells);

            return $@"
    Dim instrSheet As Worksheet
    Set instrSheet = Sheets(""{ExportViewConstants.UnformattedViewInstructionsSheetName}"")

    Dim position As Range
    Set position = instrSheet.Range(instrSheet.Cells(5, 2), instrSheet.Cells(5, 3))
    
    Dim btn As Button
    Set btn = instrSheet.Buttons.Add(position.Left, position.Top, position.Width, position.Height)
    With btn
        .Name = ""RefreshDataButton""
        .Caption = ""Refresh Data""
        .OnAction = ""ThisWorkbook.RefreshButtonClick""
    End With

    Dim sheet As Worksheet
    Set sheet = Sheets(""{ExportViewConstants.UnformattedViewDataSheetName}"")

    Dim qt As QueryTable
    Set qt = sheet.QueryTables.Add(Connection:=""URL;{_dataUrl}"", Destination:=sheet.Range(""{worksheet.Dimension.Address}""))

    qt.AdjustColumnWidth = False
    qt.RefreshStyle = xlOverwriteCells
    qt.BackgroundQuery = False

    qt.Name = ""nwshp?hl=en&tab=wn""
    qt.WebPreFormattedTextToColumns = True
    qt.WebFormatting = xlWebFormattingNone
    qt.WebConsecutiveDelimitersAsOne = True
    qt.WebDisableRedirections = False
    qt.WebSingleBlockTextImport = False
    qt.WebDisableDateRecognition = False
    qt.WebSelectionType = xlEntirePage

    sheet.Range(""{worksheet.Dimension.Address}"").Font.Bold = False
";
        }

        private static string GetRefreshDataScript()
        {
            return $@"
Sub RefreshButtonClick()
    Dim sheet As Worksheet
    Set sheet = Sheets(""{ExportViewConstants.UnformattedViewDataSheetName}"")

    Dim qt As QueryTable
    Set qt = sheet.QueryTables(1)
    qt.Refresh

    MsgBox ""Data was updated""
    sheet.Activate
End Sub
";
        }

        private static string GetPivotScript(string dataDimension)
        {
            return $@"
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets.Add After:= ActiveSheet
    ActiveSheet.Name = ""{ExportViewConstants.UnformattedViewPivotSheetName}""
    Application.DisplayAlerts = True

    Dim pivotSheet As Worksheet
    Set pivotSheet = Sheets(""{ExportViewConstants.UnformattedViewPivotSheetName}"")
    
    Dim dataSheet As Worksheet
    Set dataSheet = Sheets(""{ExportViewConstants.UnformattedViewDataSheetName}"")

    Set PRange = dataSheet.Range(""{dataDimension}"").CurrentRegion

    Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:= xlDatabase, SourceData:= PRange)
    Set pTable = pvtCache.CreatePivotTable(TableDestination:= pivotSheet.Cells(3, 1), TableName:= ""PivotTable"")

    'Set pt = pivotSheet.PivotTables(1)
    'For Each pf In pt.ColumnFields
    '    pf.Orientation = xlHidden
    'Next pf
";
        }
    }
}
