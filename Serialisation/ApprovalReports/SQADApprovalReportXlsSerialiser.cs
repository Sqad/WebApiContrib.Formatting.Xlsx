using OfficeOpenXml;
using SQAD.MTNext.Business.Models.Attributes;
using SQAD.MTNext.Business.Models.Core.ApprovalReport;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using SQAD.MTNext.Business.Models.FlowChart.Enums;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace WebApiContrib.Formatting.Xlsx.src.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    public class SQADApprovalReportXlsSerialiser : IXlsxSerialiser
    {
        private static void PopulateData(SqadXlsxSheetBuilderBase sheetBuilder, DataColumnCollection columns, IEnumerable<ExcelDataRow> records)
        {
            foreach (var record in records)
            {
                var dataRow = record.GetExcelCells(columns);
                sheetBuilder.AppendRow(dataRow);
            }
        }

        private void FillDataTable(DataTable dataTable, List<ApprovalReportExportData> approvalReports)
        {
            var columns = dataTable.Columns;
            for (var i = 0; i < approvalReports.Count(); i++)
            {
                var dataRow = dataTable.NewRow();

                dataRow[columns[(int)ApprovalReportElement.ResourceSet]] = approvalReports[i].ResourceSet;
                dataRow[columns[(int)ApprovalReportElement.Country]] = approvalReports[i].Country;
                dataRow[columns[(int)ApprovalReportElement.Client]] = approvalReports[i].Client;
                dataRow[columns[(int)ApprovalReportElement.Product]] = approvalReports[i].Product;
                dataRow[columns[(int)ApprovalReportElement.NameVersion]] = approvalReports[i].NameVersion;
                dataRow[columns[(int)ApprovalReportElement.SubmittedBy]] = approvalReports[i].SubmittedBy;
                dataRow[columns[(int)ApprovalReportElement.DateSubmitted]] = approvalReports[i].DateSubmitted.ToString("MM/dd/yyyy hh:mm");
                dataRow[columns[(int)ApprovalReportElement.StatusStep]] = approvalReports[i].StatusStep;
                dataRow[columns[(int)ApprovalReportElement.DateCompleted]] = approvalReports[i].DateCompleted?.ToString("MM/dd/yyyy hh:mm");
                dataRow[columns[(int)ApprovalReportElement.Days]] = approvalReports[i].Days;
                dataRow[columns[(int)ApprovalReportElement.Action]] = approvalReports[i].Action;
                dataRow[columns[(int)ApprovalReportElement.Comments]] = approvalReports[i].Comments;
                dataRow[columns[(int)ApprovalReportElement.Currency]] = approvalReports[i].Currency;
                dataRow[columns[(int)ApprovalReportElement.GrossCost]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].GrossCost?.ToString("n2") ?? "-"}";
                dataRow[columns[(int)ApprovalReportElement.NetCost]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].NetCost?.ToString("n2") ?? "-"}";
                dataRow[columns[(int)ApprovalReportElement.WorkingCost]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].WorkingCost?.ToString("n2") ?? "-"}";
                dataRow[columns[(int)ApprovalReportElement.NonWorkingCosts]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].NonWorkingCosts?.ToString("n2") ?? "-"}";
                dataRow[columns[(int)ApprovalReportElement.Fees]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].Fees?.ToString("n2") ?? "-"}";

                dataTable.Rows.Add(dataRow);
            }
        }
        private DataTable CreateApprovalReportDataTable(List<ApprovalReportExportData> approvalReports)
        {
            var columns = typeof(ApprovalReportExportData).GetProperties()
                          .SelectMany(item => item.GetCustomAttributes(typeof(ExcelColumnAttribute), false))
                          .Select(item => (ExcelColumnAttribute)item)
                          .OrderBy(item => item.Order)
                          .Select(item => new DataColumn(item.Header));

            var dataTable = new DataTable();

            dataTable.Columns.AddRange(columns.ToArray());

            FillDataTable(dataTable, approvalReports);

            return dataTable;
        }
        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(List<ApprovalReportExportData>);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetbuilderOverride)
        {
            if (!(value is List<ApprovalReportExportData> approvalReports))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var approvalReportDataTable = CreateApprovalReportDataTable(approvalReports);
            var columns = approvalReportDataTable.Columns;
            var rows = approvalReportDataTable.Rows;
            var startDateApprovalReport = approvalReports.FirstOrDefault().StartDate;
            var endDateApprovalReport = approvalReports.FirstOrDefault().EndDate;
            var approvalType = string.Join(',', approvalReports.Select(item => item.ApprovalType).Distinct());

            var sheetBuilder = new SqadXlsxApprovalReportSheetBuilder(startHeaderIndex: 5, startDataIndex: 6,
                totalCountColumns: columns.Count, totalCountRows: rows.Count, startDateApprovalReport, endDateApprovalReport,
                approvalType);
            document.AppendSheet(sheetBuilder);

            sheetBuilder.AppendColumns(columns);
            PopulateData(sheetBuilder, columns, rows.Cast<DataRow>().Select(item => new ExcelDataRow(item)));
        }
    }
}
