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
            
            worksheet.Workbook.CodeModule.Code = "Private Sub Workbook_Open()\r\n\tMsgbox \"Welcome!\"\r\nEnd Sub";
        }
    }
}
