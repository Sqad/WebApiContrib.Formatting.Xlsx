using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using SQAD.MTNext.Business.Models.FlowChart.Enums;
using SQAD.XlsxExportView;
using SQAD.XlsxExportImport.Base.Attributes;
using SQAD.XlsxExportImport.Base.Builders;
using SQAD.XlsxExportImport.Base.Interfaces;
using SQAD.XlsxExportImport.Base.Serialization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    public class SQADApprovalReportXlsSerializer : IXlsxSerializer
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
                    dataRow[columns[(int)ApprovalReportElement.GrossCost]] = approvalReports[i].GrossCost;
                } else
                {
                    countDeletedColumns++;
                }

                if (approvalReportExportRequest.IsNetCost)
                {
                    dataRow[columns[(int)ApprovalReportElement.NetCost - countDeletedColumns]] = approvalReports[i].NetCost;
                } else
                {
                    countDeletedColumns++;
                }

                dataRow[columns[(int)ApprovalReportElement.WorkingCost - countDeletedColumns]] = approvalReports[i].WorkingCost;

                if (approvalReportExportRequest.IsIncludeNonWorking)
                {
                    dataRow[columns[(int)ApprovalReportElement.NonWorkingCosts - countDeletedColumns]] = approvalReports[i].NonWorkingCosts;
                } else
                {
                    countDeletedColumns++;
                }

                if (approvalReportExportRequest.IsIncludeFees)
                {
                    dataRow[columns[(int)ApprovalReportElement.Fees - countDeletedColumns]] = approvalReports[i].Fees;
                } else
                {
                    countDeletedColumns++;
                }

                dataRow[columns[(int)ApprovalReportElement.IsEvenGroup - countDeletedColumns]] = approvalReports[i].IsEvenGroup;
                dataRow[columns[(int)ApprovalReportElement.CurrencySymbol - countDeletedColumns]] = approvalReports[i].CurrencySymbol;

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

        public bool CanSerializeType(Type valueType, Type itemType)
        {
            return valueType == typeof(ApprovalReportExportRequestModel);
        }

        public void Serialize(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder sheetbuilderOverride)
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

            //We should minus 2 for count columns for delete IsEvenGroup and CurrencySymbol from Worksheet. Because it's flag field
            var sheetBuilder = new SqadXlsxApprovalReportSheetBuilder(startHeaderIndex: 5, startDataIndex: 6,
                totalCountColumns: columns.Count - 2, totalCountRows: rows.Count, startDateApprovalReport: startDateApprovalReport,
                endDateApprovalReport: endDateApprovalReport, approvalType: approvalType);
            document.AppendSheet(sheetBuilder);

            sheetBuilder.AppendColumns(columns);
            PopulateData(sheetBuilder, columns, rows.Cast<DataRow>().Select(item => new ExcelDataRow(item)));
        }
    }
}
