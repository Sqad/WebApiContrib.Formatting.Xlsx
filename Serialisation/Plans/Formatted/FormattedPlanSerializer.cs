using SQAD.MTNext.Business.Models.FlowChart.Plan;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;
using System;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted
{
    internal class FormattedPlanSerializer : IXlsxSerialiser
    {
        public SerializerType SerializerType => SerializerType.FormattedPlan;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(ExportPlanRequest);
        }

        public void Serialise(Type itemType,
                              object value,
                              IXlsxDocumentBuilder document,
                              string sheetName,
                              string columnPrefix,
                              SqadXlsxPlanSheetBuilder sheetBuilderOverride)
        {
            if (!(value is ExportPlanRequest exportPlanRequest))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            document.AppendSheet(new FormattedPlanSheetBuilder("Formatted Plan", exportPlanRequest));
        }
    }
}
