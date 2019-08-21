using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Formatted
{
    public class SqadSummaryPlanXlsxSerializer : IXlsxSerialiser
    {
        public SerializerType SerializerType => SerializerType.SummaryPlan;
        
        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(DataTable);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetBuilderOverride)
        {
            if (!(value is DataTable dataTable))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var dataRows = dataTable.Rows.Cast<DataRow>();
            var records = dataRows.Select(x => new FormattedExcelDataRow(x)).ToList();

            var columns = dataTable.Columns;
            columns.RemoveAt(columns.Count - 1);

            var sheetBuilder = new SqadXlsxSummaryViewSheetBuilder(records.Count(x => x.IsHeader));
            document.AppendSheet(sheetBuilder);

            sheetBuilder.AppendColumns(columns);

            PopulateData(sheetBuilder, columns, records);
        }

        private static void PopulateData(SqadXlsxSheetBuilderBase sheetBuilder, DataColumnCollection columns, IEnumerable<FormattedExcelDataRow> records)
        {
            foreach (var record in records)
            {
                var dataRow = record.GetExcelCells(columns);
                sheetBuilder.AppendRow(dataRow);
            }
        }
    }
}
