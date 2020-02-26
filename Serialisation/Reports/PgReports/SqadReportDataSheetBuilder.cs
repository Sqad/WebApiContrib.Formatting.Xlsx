using System;
using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.Core.Reports;
using SQAD.MTNext.Business.Models.Core.Reports.PgReports;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Reports.PgReports
{
    public class SqadReportDataSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly PgReportSerializeDto _exportData;

        public SqadReportDataSheetBuilder(string sheetName,
            PgReportSerializeDto exportData,
            bool isReferenceSheet = false,
            bool isPreservationSheet = false,
            bool isHidden = false,
            bool shouldAutoFit = true)
            : base(sheetName, isReferenceSheet, isPreservationSheet, isHidden, shouldAutoFit)
        {
            _exportData = exportData;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            var dataTable = _exportData.Table;

            worksheet.Cells["A1"].Value = _exportData.ReportName;
            worksheet.Cells["A2"].Value = "P&G RESTRICTED";
            worksheet.Cells["A3"].Value = "From: " + _exportData.StartDate.ToShortDateString() + "-" +
                                          _exportData.EndDate.AddDays(-1).ToShortDateString();
            worksheet.Cells["A4"].Value = "Printed: " + DateTime.Now;
            worksheet.Cells["A5"].Value = dataTable.Rows.Count + " Records";

            worksheet.Cells["A8"].LoadFromDataTable(dataTable, true);
        }
    }
}
