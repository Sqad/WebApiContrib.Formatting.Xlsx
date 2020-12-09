using System.Data;
using OfficeOpenXml;
using SQAD.XlsxExportImport.Base.Builders;

namespace SQAD.XlsxExportView.Formatted
{
    internal class SqadXlsxFormattedViewScriptsSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly string _viewLabel;
        public SqadXlsxFormattedViewScriptsSheetBuilder(string viewLabel = null)
            : base(ExportViewConstants.ScriptSheetName, shouldAutoFit: false)
        {
            _viewLabel = viewLabel;
        }

        // note: tricky solution to force Excel auto-format
        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }
            string formattedViewSheetName = string.IsNullOrEmpty(_viewLabel)
                  ? ExportViewConstants.FormattedViewSheetName : _viewLabel;
            var code = $@"
Private Sub Workbook_Open()
    Dim tmpSheet As Worksheet
    Set tmpSheet = Sheets(""{ExportViewConstants.ScriptSheetName}"")
    If tmpSheet.Visible = xlSheetVeryHidden Then
        Exit Sub
    End If
    
    Sheets(""{formattedViewSheetName}"").UsedRange.Cells.Value = Sheets(""{formattedViewSheetName}"").UsedRange.Cells.Value

    tmpSheet.Visible = xlSheetVeryHidden
End Sub";

            worksheet.Workbook.CodeModule.Code = code;
        }
    }
}