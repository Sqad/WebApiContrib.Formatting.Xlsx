using OfficeOpenXml;
using SQAD.MTNext.Business.Models.Attributes;
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

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    public class SQADBillingReportXlsSerialiser : IXlsxSerialiser
    {
        private static void PopulateData(SqadXlsxSheetBuilderBase sheetBuilder, DataColumnCollection columns, IEnumerable<ExcelDataRow> records)
        {
            foreach (var record in records)
            {
                var dataRow = record.GetExcelCells(columns);
                sheetBuilder.AppendRow(dataRow);
            }
        }

        private void FillDataTable(DataTable dataTable, BillingReportExportRequestModel approvalReportExportRequest)
        {
            var columns = dataTable.Columns;
            var approvalReports = approvalReportExportRequest.BillingReports;

            for (var i = 0; i < approvalReports.Count(); i++)
            {
                var dataRow = dataTable.NewRow();

                dataRow[columns[(int)BillingReportElement.Email]] = approvalReports[i].Email;
                dataRow[columns[(int)BillingReportElement.LastLoggedIn]] = approvalReports[i].LastLoggedIn?.ToString("MM/dd/yyyy hh:mm");
                dataRow[columns[(int)BillingReportElement.LocationName]] = approvalReports[i].LocationName;
                dataRow[columns[(int)BillingReportElement.DatabaseName]] = approvalReports[i].DatabaseName;
                dataRow[columns[(int)BillingReportElement.DatabaseServerName]] = approvalReports[i].DatabaseServerName;
                dataRow[columns[(int)BillingReportElement.AppVersion]] = approvalReports[i].AppVersion;
                dataRow[columns[(int)BillingReportElement.BillingReference]] = approvalReports[i].BillingReference;
                dataRow[columns[(int)BillingReportElement.LoginName]] = approvalReports[i].LoginName;
                dataRow[columns[(int)BillingReportElement.Tax]] = approvalReports[i].Tax;
                dataRow[columns[(int)BillingReportElement.Type]] = approvalReports[i].Type;

                dataTable.Rows.Add(dataRow);
            }
        }
        private DataTable CreateApprovalReportDataTable(BillingReportExportRequestModel billingReportExportRequest)
        {
            var columns = typeof(BillingReportExportDataModel).GetProperties()
                          .SelectMany(item => item.GetCustomAttributes(typeof(ExcelColumnAttribute), false))
                          .Select(item => (ExcelColumnAttribute)item);

            var dataTable = new DataTable();

            dataTable.Columns.AddRange(columns.Select(item => new DataColumn(item.Header)).ToArray());

            FillDataTable(dataTable, billingReportExportRequest);

            return dataTable;
        }
        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(BillingReportExportRequestModel);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetbuilderOverride)
        {
            if (!(value is BillingReportExportRequestModel approvalReportExportRequest))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var approvalReportDataTable = CreateApprovalReportDataTable(approvalReportExportRequest);
            var columns = approvalReportDataTable.Columns;
            var rows = approvalReportDataTable.Rows;

            var sheetBuilder = new SqadXlsxBillingReportSheetBuilder(startHeaderIndex: 5, startDataIndex: 6,
                totalCountColumns: columns.Count, totalCountRows: rows.Count);
            document.AppendSheet(sheetBuilder);

            sheetBuilder.AppendColumns(columns);
            PopulateData(sheetBuilder, columns, rows.Cast<DataRow>().Select(item => new ExcelDataRow(item)));
        }
    }
}
