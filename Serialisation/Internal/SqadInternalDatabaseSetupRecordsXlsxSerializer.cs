using System;
using System.Collections.Generic;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Parsing;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Result;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;
using WebApiContrib.Formatting.Xlsx.Models;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Internal
{
    //todo vv: need to define use common types (IXlsxSerialiser, IXlsxDocumentBuilder, etc.) or just create separate simple formatter
    public class SqadInternalDatabaseSetupRecordsXlsxSerializer : IXlsxSerialiser
    {
        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return typeof(IEnumerable<ExportResultItem<ExcelRowBase>>).IsAssignableFrom(valueType);
        }

        public void Serialise(Type itemType,
                              object value,
                              IXlsxDocumentBuilder document,
                              string sheetName,
                              string columnPrefix,
                              SqadXlsxPlanSheetBuilder sheetBuilderOverride)
        {
            if (!(value is IEnumerable<ExportResultItem<ExcelRowBase>> exportResults))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            document.SetTemplateInfo(new XlsxTemplateInfo("ExcelTemplates/Internal_Database_Setup.xlsx",
                                                          null)); //todo: add password protection

            document.AppendSheet(new SqadXlsxInternalDatabaseSetupReferenceSheetBuilder(exportResults));
        }
    }
}