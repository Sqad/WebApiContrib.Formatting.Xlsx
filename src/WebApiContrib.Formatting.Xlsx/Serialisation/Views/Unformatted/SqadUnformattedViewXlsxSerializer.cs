using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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

            var tables = dataSet.Tables.Cast<DataTable>().ToList();

            ProcessInstructionsSheet(document, tables);

            var dataUrl = GetDataUrl(tables);
            ProcessDataSheet(document, tables, dataUrl);
            ProcessPivotSheet(document, tables);
        }

        private static void ProcessInstructionsSheet(IXlsxDocumentBuilder document, IEnumerable<DataTable> tables)
        {
            var instructionsDataTable = tables.FirstOrDefault(x => x.TableName == "instructions");
            if (instructionsDataTable == null)
            {
                return;
            }

            var instructionsSheetBuilder = new SqadXlsxUnformattedViewInstructionsSheetBuilder();
            document.AppendSheet(instructionsSheetBuilder);
            
            AppendColumnsAndRows(instructionsSheetBuilder, instructionsDataTable);
        }

        private static void ProcessPivotSheet(IXlsxDocumentBuilder document, IEnumerable<DataTable> tables)
        {
            var pivotDataTable = tables.FirstOrDefault(x => x.TableName == "pivot");
            if (pivotDataTable == null)
            {
                return;
            }

            var pivotSheetBuilder = new SqadXlsxUnformattedViewPivotSheetBuilder();
            document.AppendSheet(pivotSheetBuilder);
        }

        private static void ProcessDataSheet(IXlsxDocumentBuilder document, IEnumerable<DataTable> tables, string dataUrl)
        {
            var dataTable = tables.FirstOrDefault(x => x.TableName == "data");
            if (dataTable == null)
            {
                return;
            }
            
            var dataSheetBuilder = new SqadXlsxUnformattedViewDataSheetBuilder(dataUrl);
            document.AppendSheet(dataSheetBuilder);
            
            AppendColumnsAndRows(dataSheetBuilder, dataTable);
        }

        private static void AppendColumnsAndRows(SqadXlsxSheetBuilderBase sheetBuilder, DataTable dataTable)
        {
            var columns = FixColumnNames(dataTable);

            sheetBuilder.AppendColumns(columns);

            var records = dataTable.Rows.Cast<DataRow>().Select(x => new ExcelDataRow(x));
            foreach (var record in records)
            {
                var row = record.GetExcelCells(columns);
                sheetBuilder.AppendRow(row);
            }
        }

        private static DataColumnCollection FixColumnNames(DataTable dataTable)
        {
            var columns = dataTable.Columns;
            foreach (DataColumn column in columns)
            {
                if (string.Compare(column.ColumnName, "nep", StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    column.ColumnName = column.ColumnName.ToUpper();
                    continue;
                }
                column.ColumnName = $"{column.ColumnName.First().ToString().ToUpper()}{column.ColumnName.Substring(1)}";
            }

            return columns;
        }

        private static string GetDataUrl(IEnumerable<DataTable> tables)
        {
            var settingsDataTable = tables.FirstOrDefault(x => x.TableName == "_settings");

            return (string) settingsDataTable?.Select("key = 'ExcelLink'").FirstOrDefault()?["value"];
        }
    }
}
