using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace WebApiContrib.Formatting.Xlsx.src.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    public class SqadXlsxApprovalReportSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly int _startHeaderIndex;
        private readonly int _startDataIndex;
        private readonly int _totalCountColumns;
        private readonly int _totalCountRows;

        public SqadXlsxApprovalReportSheetBuilder(int startHeaderIndex, int startDataIndex, int totalCountColumns, int totalCountRows) : base(ExportConstants.ApprovalReportSheetName)
        {
            _startHeaderIndex = startHeaderIndex;
            _startDataIndex = startDataIndex;
            _totalCountColumns = totalCountColumns;
            _totalCountRows = totalCountRows;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            for (int i = 0; i < _totalCountColumns; i++)
            {
                worksheet.SetValue(_startHeaderIndex, i + 1, table.Columns[i]);
                for (var j = 0; j < _totalCountRows; j++)
                {
                    worksheet.SetValue(_startHeaderIndex, i + 1, table.Rows[i][j]);
                }
            }
        }
    }
}
