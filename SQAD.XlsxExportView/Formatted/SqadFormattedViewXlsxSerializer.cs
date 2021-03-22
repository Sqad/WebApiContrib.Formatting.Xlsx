using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SQAD.XlsxExportImport.Base.Builders;
using SQAD.XlsxExportImport.Base.Interfaces;
using SQAD.XlsxExportImport.Base.Serialization;

namespace SQAD.XlsxExportView.Formatted
{
    public class SqadFormattedViewXlsxSerializer : IXlsxSerialiser
    {
        private readonly string _viewLabel;
        public SerializerType SerializerType => SerializerType.Default;

        public SqadFormattedViewXlsxSerializer(string viewLabel = null)
        {
            _viewLabel = viewLabel;
        }
        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(DataTable);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder sheetBuilderOverride)
        {
            if (!(value is DataTable dataTable))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var dataRows = dataTable.Rows.Cast<DataRow>();
            var records = dataRows.Select(x => new FormattedExcelDataRow(x)).ToList();

            var columns = dataTable.Columns;
            columns.RemoveAt(columns.Count - 1);

            var sheetBuilder = new SqadXlsxFormattedViewSheetBuilder(records.Count(x => x.IsHeader), _viewLabel);
            document.AppendSheet(sheetBuilder);

            sheetBuilder.AppendColumns(columns);

            PopulateData(sheetBuilder, columns, records);

            var scriptBuilder = new SqadXlsxFormattedViewScriptsSheetBuilder(_viewLabel);
            document.AppendSheet(scriptBuilder);
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
