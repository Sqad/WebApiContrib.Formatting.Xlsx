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
        private void FillDataTable(DataTable dataTable,List<ApprovalReportExportData> approvalReports)
        {
            for (var j = 0; j < approvalReports.Count(); j++)
            {
                dataTable.Rows[(int)ApprovalReportElement.ResourceSet][j] = approvalReports[j].ResourceSet;
                dataTable.Rows[(int)ApprovalReportElement.Country][j] = approvalReports[j].Country;
                dataTable.Rows[(int)ApprovalReportElement.Client][j] = approvalReports[j].Client;
                dataTable.Rows[(int)ApprovalReportElement.Product][j] = approvalReports[j].Product;
                dataTable.Rows[(int)ApprovalReportElement.NameVersion][j] = approvalReports[j].NameVersion;
                dataTable.Rows[(int)ApprovalReportElement.SubmittedBy][j] = approvalReports[j].SubmittedBy;
                dataTable.Rows[(int)ApprovalReportElement.DateSubmitted][j] = approvalReports[j].DateSubmitted;
                dataTable.Rows[(int)ApprovalReportElement.StatusStep][j] = approvalReports[j].StatusStep;
                dataTable.Rows[(int)ApprovalReportElement.DateCompleted][j] = approvalReports[j].DateCompleted;
                dataTable.Rows[(int)ApprovalReportElement.Days][j] = approvalReports[j].Days;
                dataTable.Rows[(int)ApprovalReportElement.Action][j] = approvalReports[j].Action;
                dataTable.Rows[(int)ApprovalReportElement.Comments][j] = approvalReports[j].Comments;
                dataTable.Rows[(int)ApprovalReportElement.Currency][j] = approvalReports[j].Currency;
                dataTable.Rows[(int)ApprovalReportElement.GrossCost][j] = approvalReports[j].GrossCost;
                dataTable.Rows[(int)ApprovalReportElement.WorkingCost][j] = approvalReports[j].WorkingCost;
                dataTable.Rows[(int)ApprovalReportElement.NonWorkingCosts][j] = approvalReports[j].NonWorkingCosts;
                dataTable.Rows[(int)ApprovalReportElement.Fees][j] = approvalReports[j].Fees;
            }
        }
        private DataTable CreateApprovalReportDataTable(List<ApprovalReportExportData> approvalReports)
        {
            var columns = typeof(ApprovalReportExportData)
                          .GetCustomAttributes(typeof(ExcelColumnAttribute), false)
                          .Select(item => (ExcelColumnAttribute)item)
                          .OrderBy(item => item.Order);

            var dataTable = new DataTable();

            dataTable.Columns.AddRange(columns.Select(item => new DataColumn(item.Header)).ToArray());
            FillDataTable(dataTable, approvalReports);

            return dataTable;
        }
        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(ApprovalReportExportData);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetbuilderOverride)
        {
            if (!(value is IEnumerable<ApprovalReportExportData> approvalReports))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var approvalReportDataTable = CreateApprovalReportDataTable(approvalReports.ToList());

            var sheetBuilder = new SqadXlsxApprovalReportSheetBuilder(startHeaderIndex: 5, startDataIndex: 6,
                totalCountColumns: approvalReportDataTable.Columns.Count, totalCountRows: approvalReportDataTable.Rows.Count);

            document.AppendSheet(sheetBuilder);
        }
    }
}
