using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SQAD.MTNext.Business.Models.Attributes;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class SQADDataTableXlsxSerialiser : IXlsxSerialiser
    {
        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(DataTable);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName)
        {
            if (!(value is DataTable dataTable))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var dataRows = dataTable.Rows.Cast<DataRow>();
            var records = dataRows.Select(x => new ExcelDataRow(x)).ToList();

            var sheetBuilder = new SqadXlsxSheetBuilder("Sheet1");
            document.AppendSheet(sheetBuilder);
            
            CreateHeader(sheetBuilder, records);


        }

        private static void CreateHeader(SqadXlsxSheetBuilder sheetBuilder, IEnumerable<ExcelDataRow> records)
        {
            var headerRecords = records.Where(x => x.IsHeader);

            var rowIndex = 0;
            foreach (var headerRecord in headerRecords)
            {
                var columnInfoCollection = new ExcelColumnInfoCollection();
                
                var columnsInfo = headerRecord.GetExcelColumnsInfo(rowIndex);
                foreach (var excelColumnInfo in columnsInfo)
                {
                    columnInfoCollection.Add(excelColumnInfo);
                }

                sheetBuilder.AppendColumnHeaderRow(columnInfoCollection);
                rowIndex++;
            }
        }

        private class ExcelDataRow
        {
            private readonly DataRow _dataRow;

            private string this[int columnIndex] => _dataRow.IsNull(columnIndex) ? null : (string)_dataRow[columnIndex];
            private string this[string columnName] => _dataRow.IsNull(columnName) ? null : (string)_dataRow[columnName];

            public ExcelDataRow(DataRow dataRow)
            {
                _dataRow = dataRow;
                IsHeader = bool.Parse(this["header"]);
            }

            public bool IsHeader { get; }

            public IEnumerable<ExcelColumnInfo> GetExcelColumnsInfo(int rowIndex)
            {
                for (var i = 0; i < _dataRow.ItemArray.Length - 1; i++)
                {
                    yield return new ExcelColumnInfo($"{rowIndex}_{i}", typeof(string), new ExcelColumnAttribute(this[i]));
                }
            }
        }
    }
}
