using System;
using System.Data;
using System.Linq;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadUnformattedViewXlsxSerializer : IXlsxSerialiser
    {
        private const string InstructionsTableName = "Instructions";
        private const string PivotTableName = "Pivot";
        private const string DataTableName = "Data";
        private const string SettingsTableName = "_settings";

        public SerializerType SerializerType => SerializerType.Default;

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(DataSet);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetBuilderOverride)
        {
            if (!(value is DataSet dataSet))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var tables = dataSet.Tables;

            ProcessInstructionsSheet(document, tables);

            var dataUrl = GetDataUrl(tables);
            ProcessDataSheet(document, tables);

            var needCreatePivotSheet = tables.Contains(PivotTableName);
            var scriptBuilder = new SqadXlsxUnformattedViewScriptSheetBuilder(dataUrl, needCreatePivotSheet);
            document.AppendSheet(scriptBuilder);
        }

        private static void ProcessInstructionsSheet(IXlsxDocumentBuilder document, DataTableCollection tables)
        {
            if (!tables.Contains(InstructionsTableName))
            {
                return;
            }

            var instructionsDataTable = tables[InstructionsTableName];

            var instructionsSheetBuilder = new SqadXlsxUnformattedViewInstructionsSheetBuilder();
            document.AppendSheet(instructionsSheetBuilder);

            AppendColumnsAndRows(instructionsSheetBuilder, instructionsDataTable);
        }

        private static void ProcessDataSheet(IXlsxDocumentBuilder document,
                                             DataTableCollection tables)
        {
            if (!tables.Contains(DataTableName))
            {
                return;
            }

            var dataTable = tables[DataTableName];

            var dataSheetBuilder = new SqadXlsxUnformattedViewDataSheetBuilder();
            document.AppendSheet(dataSheetBuilder);

            AppendColumnsAndRows(dataSheetBuilder, dataTable);
        }

        private static void AppendColumnsAndRows(SqadXlsxSheetBuilderBase sheetBuilder, DataTable dataTable)
        {
            var columns = dataTable.Columns;

            sheetBuilder.AppendColumns(columns);

            var records = dataTable.Rows.Cast<DataRow>().Select(x => new ExcelDataRow(x));
            foreach (var record in records)
            {
                var row = record.GetExcelCells(columns);
                sheetBuilder.AppendRow(row);
            }
        }

        private static string GetDataUrl(DataTableCollection tables)
        {
            if (!tables.Contains(SettingsTableName))
            {
                return null;
            }

            var settingsDataTable = tables[SettingsTableName];

            return (string)settingsDataTable.Select("key = 'ExcelLink'").FirstOrDefault()?["value"];
        }
    }
}