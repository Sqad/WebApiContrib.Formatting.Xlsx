using System;
using System.Data;
using System.Linq;
using SQAD.XlsxExportImport.Base.Builders;
using SQAD.XlsxExportImport.Base.Interfaces;
using SQAD.XlsxExportImport.Base.Serialization;
using SQAD.XlsxExportView.Unformatted.Models;

namespace SQAD.XlsxExportView.Unformatted
{
    public class SqadUnformattedViewXlsxSerializer : IXlsxSerialiser
    {
        private const string InstructionsTableName = "Instructions";
        private const string PivotTableName = "Pivot";
        private readonly string _dataTableName = ExportViewConstants.UnformattedViewDataSheetName;
        private const string SettingsTableName = "_settings";

        public SerializerType SerializerType => SerializerType.Default;

        public SqadUnformattedViewXlsxSerializer(string viewLabel = null)
        {
            _dataTableName = string.IsNullOrEmpty(viewLabel)
                ? ExportViewConstants.UnformattedViewDataSheetName : viewLabel;
        }
        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return valueType == typeof(DataSet);
        }

        public void Serialise(Type itemType,
                              object value,
                              IXlsxDocumentBuilder document,
                              string sheetName,
                              string columnPrefix,
                              XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder sheetBuilderOverride)
        {
            if (!(value is DataSet dataSet))
            {
                throw new ArgumentException($"{nameof(value)} has invalid type!");
            }

            var tables = dataSet.Tables;

            if (!string.Equals(_dataTableName, ExportViewConstants.UnformattedViewDataSheetName))
            {
                var dataTable = tables[ExportViewConstants.UnformattedViewDataSheetName];
                dataTable.TableName = _dataTableName;
            }

            ProcessInstructionsSheet(document, tables);

            var settings = GetSettings(tables);
            ProcessDataSheet(document, tables);

            var needCreatePivotSheet = tables.Contains(PivotTableName);
            var scriptBuilder = new SqadXlsxUnformattedViewScriptSheetBuilder(settings, needCreatePivotSheet, _dataTableName);
            document.AppendSheet(scriptBuilder);
        }

        private void ProcessInstructionsSheet(IXlsxDocumentBuilder document, DataTableCollection tables)
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

        private void ProcessDataSheet(IXlsxDocumentBuilder document,
                                             DataTableCollection tables)
        {
            if (!tables.Contains(_dataTableName))
            {
                return;
            }

            var dataTable = tables[_dataTableName];

            //note: dirty fix, remove dummy row for JSON deserialization
            dataTable.Rows.RemoveAt(0);

            var dataSheetBuilder = new SqadXlsxUnformattedViewDataSheetBuilder(_dataTableName);
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

        private static UnformattedExportSettings GetSettings(DataTableCollection tables)
        {
            if (!tables.Contains(SettingsTableName))
            {
                return null;
            }

            var settingsDataTable = tables[SettingsTableName];

            return new UnformattedExportSettings(settingsDataTable);
        }
    }
}