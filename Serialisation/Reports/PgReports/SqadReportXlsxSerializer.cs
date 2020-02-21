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
    public class SqadReportXlsxSerializer : IXlsxSerialiser
    {
        private readonly IExportHelpersRepository _exportHelpersRepository;
        public SqadReportXlsxSerializer(IExportHelpersRepository exportHelpersRepository)
        {
            _exportHelpersRepository = exportHelpersRepository;
        }

        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(ApprovalHistoryReportRequestModel) || 
                valueType == typeof(GlobalReachReportRequestModel);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetbuilderOverride)
        {
            if (value is ApprovalHistoryReportRequestModel exportDataApprovalHistory)
            {
                var sheetBuilder = new SqadReportDataSheetBuilder("Data", exportDataApprovalHistory);
                document.AppendSheet(sheetBuilder);
            }
            else if (value is GlobalReachReportRequestModel exportDataGlobalReach)
            {
                
            }
            else
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            
        }
    }
}
