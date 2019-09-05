using System.Data;
using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Views.Formatted
{
    internal class SqadXlsxFormattedViewScriptsSheetBuilder : SqadXlsxSheetBuilderBase
    {
        public SqadXlsxFormattedViewScriptsSheetBuilder()
            : base(ExportViewConstants.ScriptSheetName, shouldAutoFit: false)
        {
        }

        // note: tricky solution to force Excel auto-format
        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }

            var code = $@"
Private Sub Workbook_Open()
    Dim tmpSheet As Worksheet
    Set tmpSheet = Sheets(""{ExportViewConstants.ScriptSheetName}"")
    If tmpSheet.Visible = xlSheetVeryHidden Then
        Exit Sub
    End If
    
    Sheets(""{ExportViewConstants.FormattedViewSheetName}"").UsedRange.Cells.Value = Sheets(""{ExportViewConstants.FormattedViewSheetName}"").UsedRange.Cells.Value

    tmpSheet.Visible = xlSheetVeryHidden
End Sub";

            worksheet.Workbook.CodeModule.Code = code;
        }
    }
}