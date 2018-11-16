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
            if (_dataUrl == null)
            {
                return;
            }

            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }

            var range = worksheet.Dimension.Address;

            var code = $@"
Private Sub Workbook_Open()
    If ActiveWorkbook.Connections.Count = 1 Then
        Exit Sub
    End If

    Dim sheet As Worksheet
    Set sheet = Sheets(""Data"")

    With sheet.QueryTables.Add(Connection:=""URL;{_dataUrl}"", Destination:=sheet.Range(""{range}""))
        .Name = ""nwshp?hl=en&tab=wn""
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SaveData = True
        .WebPreFormattedTextToColumns = True
        .EnableRefresh = True
        .EnableEditing = True
        .WebFormatting = xlWebFormattingNone
        .AdjustColumnWidth = False
        .Refresh BackgroundQuery:=False
    End With
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
