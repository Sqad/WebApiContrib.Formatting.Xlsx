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
        private const string InstructionsDataTableName = "instructions";
        private const string PivotTableName = "pivot";
        private const string DataTableName = "data";
        private const string SettingsTableName = "_settings";

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
            var instructionsDataTable = tables.FirstOrDefault(x => x.TableName == InstructionsDataTableName);
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
            var pivotDataTable = tables.FirstOrDefault(x => x.TableName == PivotTableName);
            if (pivotDataTable == null)
            {
                return;
            }

            var pivotSheetBuilder = new SqadXlsxUnformattedViewPivotSheetBuilder();
            document.AppendSheet(pivotSheetBuilder);
        }

        private static void ProcessDataSheet(IXlsxDocumentBuilder document, IEnumerable<DataTable> tables, string dataUrl)
        {
            var dataTable = tables.FirstOrDefault(x => x.TableName == DataTableName);
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
            sheetBuilder.AppendColumns(dataTable.Columns);

            var records = dataTable.Rows.Cast<DataRow>().Select(x => new UnformattedExcelDataRow(x));
            foreach (var record in records)
            {
                var row = record.GetExcelCells(dataTable.Columns);
                sheetBuilder.AppendRow(row);
            }
        }

        private static string GetDataUrl(IEnumerable<DataTable> tables)
        {
            var settingsDataTable = tables.FirstOrDefault(x => x.TableName == SettingsTableName);
            if (settingsDataTable == null)
            {
                return null;
            }

            var rows = settingsDataTable.Rows;
            foreach (DataRow dataRow in rows)
            {
                var key = dataRow.IsNull("key") ? null : (string) dataRow["key"];
                if (key == "dataUrl")
                {
                    return dataRow.IsNull("value") ? null : (string) dataRow["value"];
                }
            }

            return null;
        }
    }
}
