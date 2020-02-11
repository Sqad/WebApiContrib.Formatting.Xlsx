using System;
using System.Data;
using System.Threading.Tasks;
using SQAD.MTNext.Business.Auth;
using SQAD.MTNext.Business.Models.Core.Reports;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.Services.Repositories.Export;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Reports.PgReports
{
    public class SqadPgReportsXlsxSerializer : IXlsxSerialiser
    {
        private readonly IExportHelpersRepository _exportHelpersRepository;
        public SqadPgReportsXlsxSerializer(IExportHelpersRepository exportHelpersRepository)
        {
            _exportHelpersRepository = exportHelpersRepository;
        }

        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(ApprovalHistoryReportRequestModel);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetbuilderOverride)
        {
            if (!(value is ApprovalHistoryReportRequestModel exportData))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var dataTable = CreateDataTable(exportData);
            var sheetBuilder = new SqadPgReportsDataSheetBuilder("Data", dataTable, exportData);
            document.AppendSheet(sheetBuilder);
        }

        private DataTable CreateDataTable(ApprovalHistoryReportRequestModel exportData)
        {
            var query = CreateQuery(exportData);

            var dataTable = new DataTable();

            if (!string.IsNullOrEmpty(query))
            {
                var task = Task.Run(async () => await _exportHelpersRepository.GetRecordsByQuery(query));
                dataTable = task.Result;
            }

            return dataTable;
        }

        private string CreateQuery(ApprovalHistoryReportRequestModel exportData)
        {
            var query = string.Empty;

            var typesIdsStr = string.Empty;
            if (exportData.Types.Count > 0)
            {
                typesIdsStr = string.Join(",", exportData.Types);
            }

            var clientIdsStr = string.Empty;
            if (exportData.Clients.Count > 0)
            {
                clientIdsStr = string.Join(",", exportData.Clients);
            }

            var productTableJoinString = string.Empty;
            if (exportData.Products.Count > 0)
            {
                var productIdsStr = "(" + string.Join("), (", exportData.Products) + ")";

                productTableJoinString =
                    "join(SELECT * FROM(VALUES"+ productIdsStr + ") productTable(productId)) as t2 " +
                    "on t1.ProductID = t2.productId";
            }

            query =
                "SELECT Region, Country, Brand as [Brand Name], Type, Contentname as Name,  Version, submitter as [Submitted By],";
            query +=
                " CAST(Dateinitiated as datetime) as [Submitted Date], approvalstatus as [Status], CAST(DateCompleted as datetime) as [Date Completed],";
            query += 
                " Days as [Days to Approve], memberaction as Action, comments as [Submitter Comments],";
            query +=
                " concurrer as [Concur User], ConcurMemberStatus as [Concur Status], CAST(ConcurDate as datetime) as [Concur Date], ";
            query += 
                " CAST(Duedate as datetime) as [Due Date], dayslate as [Days Late] ";
            query += 
                " FROM vwApprovalReport as t1 ";
            query +=
                !string.IsNullOrEmpty(productTableJoinString) ? productTableJoinString  : "";
            query +=
                " WHERE  t1.dateinitiated between " + "'"+ exportData.StartDate.ToString("yyyy-MM-dd") + "' and '" + exportData.EndDate.ToString("yyyy-MM-dd") + "'";
            query +=
                !string.IsNullOrEmpty(typesIdsStr) ? " and t1.contenttype IN (" + typesIdsStr + ") " : "";
            query +=
                !string.IsNullOrEmpty(clientIdsStr) ? " and t1.clientID IN (" + clientIdsStr + ") " : "";
            query +=
                " ORDER BY country, brand, ContentName, DateInitiated DESC, MemberActionDate DESC ";

            return query;

        }
    }
}
