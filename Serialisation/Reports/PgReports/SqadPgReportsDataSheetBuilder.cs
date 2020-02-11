using System;
using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.Core.Reports;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Reports.PgReports
{
    public class SqadPgReportsDataSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly DataTable _dataTable;
        private readonly ApprovalHistoryReportRequestModel _exportData;

        public SqadPgReportsDataSheetBuilder(string sheetName,
            DataTable dataTable,
            ApprovalHistoryReportRequestModel exportData,
            bool isReferenceSheet = false,
            bool isPreservationSheet = false,
            bool isHidden = false,
            bool shouldAutoFit = true)
            : base(sheetName, isReferenceSheet, isPreservationSheet, isHidden, shouldAutoFit)
        {
            _dataTable = dataTable;
            _exportData = exportData;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            worksheet.Cells["A1"].Value = "APPROVAL HISTORY";
            worksheet.Cells["A2"].Value = "P&G RESTRICTED";
            worksheet.Cells["A3"].Value = "From: " + _exportData.StartDate.ToShortDateString() + "-" +
                                          _exportData.EndDate.ToShortDateString();
            worksheet.Cells["A4"].Value = "Printed: " + DateTime.Now;
            worksheet.Cells["A5"].Value = _dataTable.Rows.Count + " Records";

            worksheet.Cells["A8"].LoadFromDataTable(_dataTable, true);
        }
    }
}
