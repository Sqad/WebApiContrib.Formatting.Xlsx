using System;
using System.Data;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadUnformattedViewXlsxSerializer: IXlsxSerialiser
    {
        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(DataSet);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName)
        {
            if (!(value is DataSet dataSet))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            /*var dataRows = dataTable.Rows.Cast<DataRow>();
            var records = dataRows.Select(x => new ExcelDataRow(x)).ToList();

            var columns = dataTable.Columns;
            columns.RemoveAt(columns.Count - 1);
            
            var dataSheetBuilder = new SqadXlsxUnformattedViewDataSheetBuilder("Data");*/

        }
    }
}
