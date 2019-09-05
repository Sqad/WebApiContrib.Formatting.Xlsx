using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Actuals
{
    public class SqadActualSheetBuilder : SqadXlsxSheetBuilderBase
    {
        public SqadActualSheetBuilder(string sheetName, bool isReferenceSheet = false, bool isPreservationSheet = false, bool isHidden = false)
                : base(sheetName, isReferenceSheet, isPreservationSheet, isHidden)
        {
            //_sheetCodeColumnStatements = new List<string>();
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {

            var headerRow = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns];
            headerRow.Style.Font.Size = 20;
            headerRow.Style.Font.Bold = true;

            var subHeaderRow = worksheet.Cells[2, 1, 2, worksheet.Dimension.Columns];
            subHeaderRow.Style.Font.Size = 14;
            subHeaderRow.Style.Font.Bold = true;

            worksheet.Column(1).Width = 20;
        }

    }
}
