using System.Data;
using OfficeOpenXml;
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

            FormatDateTimeColumn(worksheet, "StartDate");
            FormatDateTimeColumn(worksheet, "EndDate");
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
            
            const string cellsName = "nwshp?hl_en_tab_wn";
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
    Set qt = sheet.QueryTables.Add(Connection:=""URL;{_dataUrl}"", Destination:=sheet.Range(""{cellsName}""))

    qt.Name=""nwshp?hl=en&tab=wn""
    qt.BackgroundQuery = False
    qt.RefreshStyle = xlOverwriteCells
    qt.WebPreFormattedTextToColumns = True
    qt.WebFormatting = xlWebFormattingNone
    qt.AdjustColumnWidth = False
    qt.WebConsecutiveDelimitersAsOne = True
    qt.WebDisableRedirections = False
    qt.WebSingleBlockTextImport = False
    qt.WebDisableDateRecognition = False
    qt.WebSelectionType = xlEntirePage
End Sub
";

            worksheet.Workbook.CodeModule.Code = code;
        }

        private void FormatDateTimeColumn(ExcelWorksheet worksheet, string columnName)
        {
            var columnIndex = CurrentTable.Columns.IndexOf(columnName) + 1;
            var column = worksheet.Cells[2, columnIndex, worksheet.Dimension.Rows, columnIndex];
            column.Style.Numberformat.Format = "dd-mm-yy";
        }
    }
}
