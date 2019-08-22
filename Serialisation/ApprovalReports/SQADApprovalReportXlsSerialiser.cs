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
        private IEnumerable<ExcelColumnAttribute> ApplyCollumnFilters(IEnumerable<ExcelColumnAttribute> columns, bool isGrossCost, bool isNetCost, bool isIncludeNonWorking, bool isIncludeFees)
        {
            if (!isGrossCost)
            {
                columns = columns.Where(item => item.Order != (int)ApprovalReportElement.GrossCost);
            }
            if (!isNetCost)
            {
                columns = columns.Where(item => item.Order != (int)ApprovalReportElement.NetCost);
            }
            if (!isIncludeNonWorking)
            {
                columns = columns.Where(item => item.Order != (int)ApprovalReportElement.NonWorkingCosts);
            }
            if (!isIncludeFees)
            {
                columns = columns.Where(item => item.Order != (int)ApprovalReportElement.Fees);
            }

            return columns.OrderBy(item => item.Order);
        }
        private void FillDataTable(DataTable dataTable, ApprovalReportExportRequestModel approvalReportExportRequest)
        {
            var columns = dataTable.Columns;
            var approvalReports = approvalReportExportRequest.ApprovalReports;

            for (var i = 0; i < approvalReports.Count(); i++)
            {
                var dataRow = dataTable.NewRow();
                var countDeletedColumns = 0;

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

                if (approvalReportExportRequest.IsGrossCost)
                {
                    dataRow[columns[(int)ApprovalReportElement.GrossCost]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].GrossCost?.ToString("n2") ?? "-"}";
                }
                else
                {
                    countDeletedColumns++;
                }

                if (approvalReportExportRequest.IsNetCost)
                {
                    dataRow[columns[(int)ApprovalReportElement.NetCost - countDeletedColumns]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].NetCost?.ToString("n2") ?? "-"}";
                }
                else
                {
                    countDeletedColumns++;
                }

                dataRow[columns[(int)ApprovalReportElement.WorkingCost - countDeletedColumns]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].WorkingCost?.ToString("n2") ?? "-"}";

                if (approvalReportExportRequest.IsIncludeNonWorking)
                {
                    dataRow[columns[(int)ApprovalReportElement.NonWorkingCosts - countDeletedColumns]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].NonWorkingCosts?.ToString("n2") ?? "-"}";
                }
                else
                {
                    countDeletedColumns++;
                }

                if (approvalReportExportRequest.IsIncludeFees)
                {
                    dataRow[columns[(int)ApprovalReportElement.Fees - countDeletedColumns]] = $"{approvalReports[i].CurrencySymbol} {approvalReports[i].Fees?.ToString("n2") ?? "-"}";
                }
                dataTable.Rows.Add(dataRow);
            }
        }
        private DataTable CreateApprovalReportDataTable(ApprovalReportExportRequestModel approvalReportExportRequest)
        {
            var columns = typeof(ApprovalReportExportDataModel).GetProperties()
                          .SelectMany(item => item.GetCustomAttributes(typeof(ExcelColumnAttribute), false))
                          .Select(item => (ExcelColumnAttribute)item);

            columns = ApplyCollumnFilters(columns: columns, isGrossCost: approvalReportExportRequest.IsGrossCost,
                isNetCost: approvalReportExportRequest.IsNetCost, isIncludeNonWorking: approvalReportExportRequest.IsIncludeNonWorking,
                isIncludeFees: approvalReportExportRequest.IsIncludeFees);

            var dataTable = new DataTable();

            dataTable.Columns.AddRange(columns.Select(item => new DataColumn(item.Header)).ToArray());

            FillDataTable(dataTable, approvalReportExportRequest);

            return dataTable;
        }
        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(ApprovalReportExportRequestModel);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetbuilderOverride)
        {
            if (!(value is ApprovalReportExportRequestModel approvalReportExportRequest))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var approvalReportDataTable = CreateApprovalReportDataTable(approvalReportExportRequest);
            var columns = approvalReportDataTable.Columns;
            var rows = approvalReportDataTable.Rows;
            var startDateApprovalReport = approvalReportExportRequest.StartDate;
            var endDateApprovalReport = approvalReportExportRequest.EndDate;
            var approvalType = string.Join(',', approvalReportExportRequest.ApprovalReports.Select(item => item.ApprovalType).Distinct());

            var sheetBuilder = new SqadXlsxApprovalReportSheetBuilder(startHeaderIndex: 5, startDataIndex: 6,
                totalCountColumns: columns.Count, totalCountRows: rows.Count, startDateApprovalReport, endDateApprovalReport,
                approvalType);
            document.AppendSheet(sheetBuilder);

            sheetBuilder.AppendColumns(columns);
            PopulateData(sheetBuilder, columns, rows.Cast<DataRow>().Select(item => new ExcelDataRow(item)));
        }
    }
}
