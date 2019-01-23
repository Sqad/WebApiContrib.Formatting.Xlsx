using System;
using System.Data;
using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewDataSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly string _dataUrl;

        public SqadXlsxUnformattedViewDataSheetBuilder(string dataUrl)
            : base(ExportViewConstants.UnformattedViewDataSheetName, shouldAutoFit: false)
        {
            _dataUrl = dataUrl;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            WorksheetDataHelper.FillData(worksheet, table, true);

            FormatDateTimeColumns(worksheet);
        }

        protected override void PostCompileActions(ExcelWorksheet worksheet)
        {
            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }

            if (_dataUrl == null)
            {
                return;
            }
            
            const string cellsName = "nwshp_hl_en_tab_wn";
            var cells = worksheet.Cells[worksheet.Dimension.Address];
            worksheet.Names.Add(cellsName, cells);

            var code = $@"
Private Sub Workbook_Open()
    Dim sheet As Worksheet
    Set sheet = Sheets(""{ExportViewConstants.UnformattedViewDataSheetName}"")
    
    If sheet.QueryTables.Count <> 0 Then
        Exit Sub
    End If

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
End Sub
";

            worksheet.Workbook.CodeModule.Code = code;
        }
        
        private void FormatDateTimeColumns(ExcelWorksheet worksheet)
        {
            for (var i = 0; i < CurrentTable.Columns.Count; i++)
            {
                if (!(((ExcelCell) CurrentTable.Rows[0][i]).CellValue is DateTime))
                {
                    continue;
                }

                var column = worksheet.Cells[2, i + 1, worksheet.Dimension.Rows, i + 1];
                column.Style.Numberformat.Format = "dd-mm-yy";
            }
        }
    }
}
