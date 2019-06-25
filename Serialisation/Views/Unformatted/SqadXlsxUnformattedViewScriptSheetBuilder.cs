using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System.Data;
using System.Linq;
using WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted.Models;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    internal class SqadXlsxUnformattedViewScriptSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private const string RefreshDataQueryTableName = "RefreshDataQueryTable";

        private readonly UnformattedExportSettings _settings;
        private readonly bool _needCreatePivotSheet;

        public SqadXlsxUnformattedViewScriptSheetBuilder(UnformattedExportSettings settings, bool needCreatePivotSheet)
            : base(ExportViewConstants.UnformattedViewScriptSheetName, shouldAutoFit: false)
        {
            _settings = settings;
            _needCreatePivotSheet = needCreatePivotSheet;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }

            var dataConstantsScript = string.Empty;
            var dataScript = string.Empty;
            var refreshDataScript = string.Empty;

            var dataSheet = worksheet.Workbook
                                     .Worksheets
                                     .First(x => x.Name == ExportViewConstants.UnformattedViewDataSheetName);

            if (_settings != null)
            {
                dataConstantsScript = GetDataConstantsScript();
                dataScript = GetDataScript(dataSheet);
                refreshDataScript = _settings.UseNewVersion ? GetNewRefreshDataScript() : GetOldRefreshDataScript();
            }

            var pivotScript = string.Empty;
            if (_needCreatePivotSheet)
            {
                pivotScript = GetPivotScript(dataSheet.Dimension.ToString());
            }

            var code = $@"
{dataConstantsScript}

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

        private string GetDataConstantsScript()
        {
            var constants = $@"
    Private Const ExportUrl = ""{_settings.ExcelLink}""
";

            if (!_settings.UseEmbeddedLogin)
            {
                constants = $@"
{constants}
    Private Const TokenPageLink = ""{_settings.TokenPageLink}""
    Private Const LoginPageLink = ""{_settings.LoginPageLink}""
";
            }

            return constants;
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
    Set qt = sheet.QueryTables.Add(Connection:=""URL;"" & ExportUrl, Destination:=sheet.Range(""{worksheet.Dimension.Address}""))

    qt.AdjustColumnWidth = False
    qt.RefreshStyle = xlOverwriteCells
    qt.BackgroundQuery = False

    qt.Name = ""{RefreshDataQueryTableName}""
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

        private static string GetOldRefreshDataScript()
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

        private string GetNewRefreshDataScript()
        {
            return $@"
Sub RefreshButtonClick()
    On Error GoTo error_handler

    Dim ieObject As Object
    Set ieObject = CreateObject(""InternetExplorer.Application"")
    
    ieObject.Visible = True
    ieObject.Navigate ""{_settings.TokenPageLink}""

    Dim shell As Object
    Dim eachIE As Object
    Do
        Set shell = CreateObject(""Shell.Application"")
        For Each eachIE In shell.Windows
            If InStr(1, eachIE.locationurl, ""{_settings.LoginPageLink}"") Or InStr(1, eachIE.locationurl, ""{_settings.TokenPageLink}"") Then
                Set ieObject = eachIE
                Set eachIE = Nothing
                Set shell = Nothing
            Exit Do
            End If
        Next eachIE
    Loop
    
    Do Until ieObject.ReadyState = 4 And ieObject.LocationName = ""Excel Data Exports"": DoEvents: Loop
    
    Dim token As String
    token = ieObject.Document.getElementById(""token"").Value
    
    ieObject.Quit
    Set ieObject = Nothing
    

    Dim sheet As Worksheet
    Set sheet = Sheets(""{ExportViewConstants.UnformattedViewDataSheetName}"")
    
    Dim qt As QueryTable
    Set qt = sheet.QueryTables(""{RefreshDataQueryTableName}"")
    
    qt.Connection = ""URL;"" & ExportUrl & ""&userToken="" & token
    qt.Refresh
    
    Exit Sub
error_handler:
    MsgBox (""Error while data exports"")
    
    If Not ieObject Is Nothing Then
        ieObject.Quit
        Set ieObject = Nothing
    End If
    If Not shell Is Nothing Then
        Set shell = Nothing
    End If
    If Not eachIE Is Nothing Then
        Set eachIE = Nothing
    End If
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