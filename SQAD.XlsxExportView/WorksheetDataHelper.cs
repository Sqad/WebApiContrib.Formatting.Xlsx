﻿using System.Data;
using OfficeOpenXml;
using SQAD.XlsxExportImport.Base.Models;

namespace SQAD.XlsxExportView
{
    internal static class WorksheetDataHelper
    {
        public static void FillData(ExcelWorksheet worksheet, DataTable table, bool includeColumnsRow)
        {
            if (includeColumnsRow)
            {
                for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    var columnName = table.Columns[columnIndex].ColumnName;

                    var cell = worksheet.Cells[1, columnIndex + 1];
                    cell.Value = columnName;
                }
            }

            for (var y = 0; y < table.Rows.Count; y++)
            {
                var dataRow = table.Rows[y];

                for (var x = 0; x < table.Columns.Count; x++)
                {
                    var column = table.Columns[x];
                    var value = dataRow[column.ColumnName];

                    var cell = worksheet.Cells[y + (includeColumnsRow ? 2 : 1), x + 1];

                    cell.Value = ((ExcelCell)value).CellValue;
                }
            }
        }
    }
}
