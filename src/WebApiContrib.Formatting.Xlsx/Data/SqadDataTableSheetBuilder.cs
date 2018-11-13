using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Data
{
    public class SqadDataTableSheetBuilder
    {
        private const int FirstColumnsToIgnore = 3;

        private readonly ExcelWorksheet _worksheet;
        private int _currentRowIndex = 1;

        public SqadDataTableSheetBuilder(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public void BuildSheet(DataTable dataTable)
        {
            var records = dataTable.Rows.Cast<DataRow>().Select(x => new ExcelDataRow(x)).ToList();
            CreateHeader(records);
            CreateData(records);

            FormatHeader(records);

            var cells = _worksheet.Cells[_worksheet.Dimension.Address];
            cells.AutoFitColumns();
        }

        private void FormatHeader(IEnumerable<ExcelDataRow> records)
        {
            var headerRecords = records.Where(x => x.IsHeader).ToList();

            int recordIndex = 1;
            foreach (var record in headerRecords)
            {
                var initialColumn = FirstColumnsToIgnore;
                var endColumn = initialColumn;

                var values = record.GetValues();
                for (; endColumn < values.Length; endColumn++)
                {
                    if (values[initialColumn] == null && values[endColumn] != null)
                    {
                        var cells = _worksheet.Cells[recordIndex, initialColumn + 1, recordIndex, endColumn + 1];
                        cells.Merge = true;
                        cells.Value = values[endColumn];

                        cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        initialColumn = endColumn + 2;
                        endColumn = endColumn + 2;
                    }
                }

                recordIndex++;

                if (recordIndex == headerRecords.Count)
                {
                    break;
                }
            }
        }

        private void CreateHeader(IEnumerable<ExcelDataRow> records)
        {
            var headerRecords = records.Where(x => x.IsHeader).ToList();

            foreach (var headerRecord in headerRecords)
            {
                var values = headerRecord.GetValues();

                for (var i = 0; i < values.Length; i++)
                {
                    var cell = _worksheet.Cells[_currentRowIndex, i + 1];
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(212, 227, 244));
                    cell.Value = values[i];
                }

                _currentRowIndex++;
            }
        }

        private void CreateData(IEnumerable<ExcelDataRow> records)
        {
            var dataRecords = records.Where(x => !x.IsHeader);
            foreach (var record in dataRecords)
            {
                var values = record.GetValues();

                for (var i = 0; i < values.Length; i++)
                {
                    var cell = _worksheet.Cells[_currentRowIndex, i + 1];
                    cell.Value = values[i];
                }

                _currentRowIndex++;
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

            public object[] GetValues()
            {
                var result = new List<object>();
                for (var i = 0; i < _dataRow.ItemArray.Length - 1; i++)
                {
                    result.Add(this[i]);
                }

                return result.ToArray();
            }
        }
    }
}
