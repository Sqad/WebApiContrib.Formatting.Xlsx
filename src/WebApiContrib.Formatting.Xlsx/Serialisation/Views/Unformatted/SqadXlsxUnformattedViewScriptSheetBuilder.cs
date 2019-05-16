using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System.Data;
using System.Linq;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewScriptSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly string _dataUrl;

        public SqadXlsxUnformattedViewScriptSheetBuilder(string dataUrl)
            :base(ExportViewConstants.UnformattedViewScriptSheetName, shouldAutoFit: false)
        {
            _dataUrl = dataUrl;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }

            var dataScript = string.Empty;
            var dataSheet =
                worksheet.Workbook.Worksheets.FirstOrDefault(x => x.Name == ExportViewConstants
                                                                      .UnformattedViewDataSheetName);
            if (_dataUrl != null)
            {
                dataScript = GetDataScript(dataSheet);
            }

            var pivotScript = string.Empty;
            var pivotSheet =
                worksheet.Workbook.Worksheets.FirstOrDefault(x => x.Name == ExportViewConstants
                                                                      .UnformattedViewPivotSheetName);
            //if (pivotSheet != null && pivotSheet.PivotTables.First().Fields.Any(x => x.Name == "Week"))
            //{
                pivotScript = GetPivotScript(dataSheet.Dimension.ToString());
            //}

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
";
           worksheet.Workbook.CodeModule.Code = code;
        }

        private string GetDataScript(ExcelWorksheet worksheet)
        {
            const string cellsName = "nwshp_hl_en_tab_wn";
            var cells = worksheet.Cells[worksheet.Dimension.Address];
            worksheet.Names.Add(cellsName, cells);

            return $@"
    Dim sheet As Worksheet
    Set sheet = Sheets(""{ExportViewConstants.UnformattedViewDataSheetName}"")
    
'    If sheet.QueryTables.Count <> 0 Then
'        Exit Sub
'    End If

    Dim qt As QueryTable
    Set qt = sheet.QueryTables.Add(Connection:=""URL;{_dataUrl}"", Destination:=sheet.Range(""{worksheet.Dimension.Address}""))

    qt.AdjustColumnWidth = False
    qt.RefreshStyle = xlOverwriteCells
    qt.BackgroundQuery = False

    Dim urlConnection As Variant
    urlConnection = qt.Connection
    
    Set rng = sheet.Range(""{worksheet.Dimension.Address}"")
    Set adoRecordset = CreateObject(""ADODB.Recordset"")
    Set xlXML = CreateObject(""MSXML2.DOMDocument"")
    xlXML.LoadXML rng.Value(xlRangeValueMSPersistXML)
    adoRecordset.Open xlXML

    Set qt.Recordset = adoRecordset
    qt.Refresh

    qt.Name = ""nwshp?hl=en&tab=wn""
    qt.Connection = urlConnection
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

        private string GetPivotScript(string dataDimention)
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

            Set PRange = dataSheet.Range(""{dataDimention}"").CurrentRegion

            Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:= xlDatabase, SourceData:= PRange)
            Set pTable = pvtCache.CreatePivotTable(TableDestination:= pivotSheet.Cells(3, 1), TableName:= ""PivotTable"")

            'Set pt = pivotSheet.PivotTables(1)
             '   For Each pf In pt.ColumnFields
              '      pf.Orientation = xlHidden
              '  Next pf
";
        }
    }
}
