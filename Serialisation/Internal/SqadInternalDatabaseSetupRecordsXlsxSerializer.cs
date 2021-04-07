using System;
using System.Collections.Generic;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Parsing.Base;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Result;
using SQAD.XlsxExportImport.Base.Builders;
using SQAD.XlsxExportImport.Base.Interfaces;
using SQAD.XlsxExportImport.Base.Models;
using SQAD.XlsxExportImport.Base.Serialization;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Internal
{
    //todo vv: need to define use common types (IXlsxSerializer, IXlsxDocumentBuilder, etc.) or just create separate simple formatter
    public class SqadInternalDatabaseSetupRecordsXlsxSerializer : IXlsxSerializer
    {
        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerializeType(Type valueType, Type itemType)
        {
            return typeof(IEnumerable<ExportResultItem<ExcelRowBase>>).IsAssignableFrom(valueType);
        }

        public void Serialize(Type itemType,
                              object value,
                              IXlsxDocumentBuilder document,
                              string sheetName,
                              string columnPrefix,
                              SqadXlsxSheetBuilder sheetBuilderOverride)
        {
            if (!(value is IEnumerable<ExportResultItem<ExcelRowBase>> exportResults))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            document.SetTemplateInfo(new XlsxTemplateInfo("ExcelTemplates/Internal_Database_Setup.xlsm",
                                                          null)); //todo: add password protection

            document.AppendSheet(new SqadXlsxInternalDatabaseSetupReferenceSheetBuilder(exportResults));
        }
    }
}