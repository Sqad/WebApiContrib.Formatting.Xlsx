using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Helpers;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Formatted
{
    public class SqadXlsxSummaryViewSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly NeutralColorGenerator _neutralColorGenerator = new NeutralColorGenerator();
        private readonly IDictionary<int, HashSet<int>> _cellsWithData = new Dictionary<int, HashSet<int>>();
        private readonly ICollection<int> _totalColumnIndexes = new List<int>();
        private readonly IDictionary<string, IDictionary<string, IDictionary<string, long?>>> _calculatedTotals;
        private readonly int _headerRowsCount;

        private bool _isMeasureColumnExists;
        private int _measuresCount;
        private int _leftPaneWidth;

        public SqadXlsxSummaryViewSheetBuilder(int headerRowsCount)
            : base(ExportViewConstants.FormattedViewSheetName)
        {
            _calculatedTotals = new Dictionary<string, IDictionary<string, IDictionary<string, long?>>>();
            _headerRowsCount = headerRowsCount;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            WorksheetDataHelper.FillData(worksheet, table, false);

            ObtainLeftPaneWidth(worksheet);
            ObtainMeasuresCount(worksheet);
            ObtainTotalColumns(worksheet);

            FillHeaderData(worksheet);
            RemoveTotalColumns(worksheet);
            MergeHeaderCells(worksheet);
            AppendCalculatedTotalColumns(worksheet);

            WorksheetHelpers.FormatRows(worksheet, _headerRowsCount + 1, _leftPaneWidth);
            WorksheetHelpers.FormatDataRows(worksheet, _headerRowsCount + 1, _totalColumnIndexes, _leftPaneWidth + 1);
            WorksheetHelpers.FormatHeader(worksheet, _headerRowsCount, _totalColumnIndexes);

            FormatSummaryRows(worksheet);
        }

        private void FillHeaderData(ExcelWorksheet worksheet)
        {
            for (var rowIndex = 1; rowIndex <= _headerRowsCount - 1; rowIndex++)
            {
                var startColumnIndex = _leftPaneWidth + 1;
                for (var endColumnIndex = startColumnIndex;
                     endColumnIndex <= worksheet.Dimension.Columns;
                     endColumnIndex++)
                {
                    var initialCell = worksheet.Cells[rowIndex, startColumnIndex];
                    if (initialCell.Value != null)
                    {
                        startColumnIndex++;
                        endColumnIndex = startColumnIndex;

                        continue;
                    }

                    var endCell = worksheet.Cells[rowIndex, endColumnIndex];
                    if (endCell.Value == null)
                    {
                        continue;
                    }

                    var cells = worksheet.Cells[rowIndex, startColumnIndex, rowIndex, endColumnIndex];
                    cells.Value = endCell.Value;

                    endColumnIndex += 2;
                    startColumnIndex = endColumnIndex;
                }
            }
        }

        protected override void PostCompileActions(ExcelWorksheet worksheet)
        {
            ProcessGroups(worksheet, _headerRowsCount + 1, worksheet.Dimension.Rows, 0);
        }

        private void ObtainLeftPaneWidth(ExcelWorksheet worksheet)
        {
            _leftPaneWidth = 3;
            for (var columnIndex = 0; columnIndex < worksheet.Dimension.Columns; columnIndex++)
            {
                var cell = worksheet.Cells[_headerRowsCount, columnIndex + 1];
                if (cell.Value == null)
                {
                    continue;
                }

                _leftPaneWidth = columnIndex;
                break;
            }
        }

        private void ObtainMeasuresCount(ExcelWorksheet worksheet)
        {
            var measuresCount = 0;
            var isStartFound = false;
            for (var rowIndex = _headerRowsCount + 1; rowIndex < worksheet.Dimension.Rows; rowIndex++)
            {
                if (WorksheetHelpers.IsDataRow(worksheet, rowIndex) && !isStartFound)
                {
                    isStartFound = true;
                    measuresCount++;
                    continue;
                }

                if (!isStartFound)
                {
                    continue;
                }

                if (WorksheetHelpers.IsDataRow(worksheet, rowIndex)
                    || WorksheetHelpers.IsGroupRow(worksheet, rowIndex, _totalColumnIndexes)
                    || WorksheetHelpers.IsTotalRow(worksheet, rowIndex, _leftPaneWidth + 1))
                {
                    _measuresCount = measuresCount;
                    break;
                }

                measuresCount++;
            }

            _isMeasureColumnExists = _leftPaneWidth >= WorksheetHelpers.RowMeasureColumnIndex;
        }

        private void ObtainTotalColumns(ExcelWorksheet sheet)
        {
            for (var cellIndex = _leftPaneWidth + 1; cellIndex <= sheet.Dimension.Columns; cellIndex++)
            {
                for (var rowIndex = 1; rowIndex <= _headerRowsCount; rowIndex++)
                {
                    var headerCell = sheet.Cells[rowIndex, cellIndex];
                    if (!(headerCell.Value is string headerValue) ||
                        !headerValue.Contains(WorksheetHelpers.TotalRowIndicator))
                    {
                        continue;
                    }

                    _totalColumnIndexes.Add(cellIndex);
                    break;
                }
            }
        }

        private void MergeHeaderCells(ExcelWorksheet worksheet)
        {
            var mergedCells = worksheet.Cells[1, 1, _headerRowsCount, _leftPaneWidth];
            mergedCells.Merge = true;
            WorksheetHelpers.SetBordersToCells(mergedCells);

            for (var rowIndex = 1; rowIndex <= _headerRowsCount - 1; rowIndex++)
            {
                var startColumnIndex = _leftPaneWidth + 1;
                var currentCellValue = worksheet.Cells[rowIndex, startColumnIndex].Value;

                for (var endColumnIndex = startColumnIndex + 1;
                     endColumnIndex <= worksheet.Dimension.Columns + 1;
                     endColumnIndex++)
                {
                    var endCell = worksheet.Cells[rowIndex, endColumnIndex];
                    if (endCell.Value == currentCellValue)
                    {
                        continue;
                    }

                    var cellBeforeEnd = worksheet.Cells[rowIndex, endColumnIndex - 1];

                    var horizontalCellsToMerge =
                        worksheet.Cells[rowIndex, startColumnIndex, rowIndex, endColumnIndex - 1];
                    horizontalCellsToMerge.Value = cellBeforeEnd.Value;
                    horizontalCellsToMerge.Merge = true;
                    WorksheetHelpers.SetBordersToCells(horizontalCellsToMerge);

                    startColumnIndex = endColumnIndex;
                    currentCellValue = worksheet.Cells[rowIndex, startColumnIndex].Value;
                    endColumnIndex++;
                }
            }

            for (var columnIndex = _leftPaneWidth + 1; columnIndex <= worksheet.Dimension.Columns; columnIndex++)
            {
                WorksheetHelpers.SetBordersToCells(worksheet.Cells[_headerRowsCount, columnIndex]);
            }
        }

        private void RemoveTotalColumns(ExcelWorksheet sheet)
        {
            if (_headerRowsCount < 2)
            {
                for (var columnIndex = _leftPaneWidth + 1; columnIndex <= sheet.Dimension.Columns; columnIndex++)
                {
                    SaveTotalColumn(sheet, columnIndex, true);
                }

                return;
            }

            var deletedColumnsCount = 0;
            foreach (var totalColumnIndex in _totalColumnIndexes.OrderBy(x => x))
            {
                var fixedTotalColumnIndex = totalColumnIndex - deletedColumnsCount;

                var firstCell = sheet.Cells[1, fixedTotalColumnIndex];
                if (firstCell.Value is string value
                    && value.Contains(WorksheetHelpers.TotalRowIndicator))
                {
                    SaveTotalColumn(sheet, fixedTotalColumnIndex);
                }

                sheet.DeleteColumn(fixedTotalColumnIndex);
                deletedColumnsCount++;
            }

            _totalColumnIndexes.Clear();
        }

        private void SaveTotalColumn(ExcelWorksheet sheet, int columnIndex, bool appendTotalIndicator = false)
        {
            var totalColumnName = sheet.Cells[1, columnIndex].Value.ToString();
            if (appendTotalIndicator)
            {
                totalColumnName = $"{totalColumnName} {WorksheetHelpers.TotalRowIndicator}";
            }

            _calculatedTotals.Add(totalColumnName, new Dictionary<string, IDictionary<string, long?>>());

            var uniqueRowIdBuilder = new List<string>();
            for (var rowIndex = _headerRowsCount + 1; rowIndex <= sheet.Dimension.Rows; rowIndex++)
            {
                var nameCellValue = sheet.Cells[rowIndex, WorksheetHelpers.RowNameColumnIndex].Value as string;
                if (WorksheetHelpers.IsGroupRow(sheet, rowIndex, _totalColumnIndexes) &&
                    !WorksheetHelpers.IsTotalRow(sheet, rowIndex, _leftPaneWidth + 1))
                {
                    uniqueRowIdBuilder.Add(nameCellValue);

                    continue;
                }

                if (WorksheetHelpers.IsTotalRow(sheet, rowIndex, _leftPaneWidth + 1))
                {
                    var groupRowKey = string.Join("-", uniqueRowIdBuilder);

                    var totalMeasures = GetMeasures(sheet, rowIndex, columnIndex);
                    foreach (var measure in totalMeasures)
                    {
                        if (!_calculatedTotals[totalColumnName].ContainsKey(measure.Key))
                        {
                            _calculatedTotals[totalColumnName].Add(measure.Key, new Dictionary<string, long?>());
                        }

                        var totalRowKey = $"{groupRowKey}-{nameCellValue}-{measure.Key}";

                        _calculatedTotals[totalColumnName][measure.Key].Add(groupRowKey, measure.Value);
                        _calculatedTotals[totalColumnName][measure.Key].Add(totalRowKey, measure.Value);
                    }

                    if (uniqueRowIdBuilder.Any())
                    {
                        uniqueRowIdBuilder.RemoveAt(uniqueRowIdBuilder.Count - 1);
                    }

                    rowIndex += _measuresCount - 1;
                    continue;
                }

                if (!WorksheetHelpers.IsDataRow(sheet, rowIndex))
                {
                    continue;
                }

                var rowKeyPrefix = string.Join("-", uniqueRowIdBuilder);
                var rowKey = $"{rowKeyPrefix}-{nameCellValue}";

                var measures = GetMeasures(sheet, rowIndex, columnIndex);
                foreach (var measure in measures)
                {
                    if (!_calculatedTotals[totalColumnName].ContainsKey(measure.Key))
                    {
                        _calculatedTotals[totalColumnName].Add(measure.Key, new Dictionary<string, long?>());
                    }


                    _calculatedTotals[totalColumnName][measure.Key].Add(rowKey, measure.Value);
                }

                rowIndex += _measuresCount - 1;
            }
        }

        private Dictionary<string, long?> GetMeasures(ExcelWorksheet sheet, int rowIndex, int columnIndex)
        {
            var result = new Dictionary<string, long?>();

            var uniqueMeasures = new Dictionary<string, int>();
            for (var shift = 0; shift < _measuresCount; shift++)
            {
                var measureRowIndex = rowIndex + shift;

                var measureCellValue = GetMeasureCellValue(sheet, measureRowIndex);
                if (!uniqueMeasures.ContainsKey(measureCellValue))
                {
                    uniqueMeasures.Add(measureCellValue, 0);
                }

                uniqueMeasures[measureCellValue]++;

                var totalCellValue = sheet.Cells[measureRowIndex, columnIndex].Value as long?;

                var counter = uniqueMeasures[measureCellValue];
                var uniqueMeasureValue = counter > 1 ? $"{measureCellValue} ({counter})" : measureCellValue;
                result.Add(uniqueMeasureValue, totalCellValue);
            }

            return result;
        }

        private string GetMeasureCellValue(ExcelWorksheet sheet, int rowIndex)
        {
            var measureCellValue = "";
            if (_isMeasureColumnExists)
            {
                measureCellValue = (string) sheet.Cells[rowIndex, WorksheetHelpers.RowMeasureColumnIndex].Value;
            }

            return measureCellValue;
        }

        private void AppendCalculatedTotalColumns(ExcelWorksheet sheet)
        {
            var lastColumn = sheet.Dimension.Columns + 1;
            var rowKeys = GetRowKeys(sheet);

            foreach (var yearGroup in _calculatedTotals)
            {
                var totalName = yearGroup.Key;

                foreach (var measureGroup in yearGroup.Value)
                {
                    var headerCells = sheet.Cells[1, lastColumn, _headerRowsCount, lastColumn];
                    headerCells.Value = $"{totalName} {measureGroup.Key}";

                    var dataCells = sheet.Cells[_headerRowsCount + 1, lastColumn, sheet.Dimension.Rows, lastColumn];
                    dataCells.Value = "-";

                    foreach (var values in measureGroup.Value)
                    {
                        if (!rowKeys.ContainsKey(values.Key))
                        {
                            continue;
                        }

                        var rowIndex = rowKeys[values.Key];

                        var valueCell = sheet.Cells[rowIndex, lastColumn];

                        if (values.Value.HasValue)
                        {
                            valueCell.Value = values.Value;
                        }
                    }

                    _totalColumnIndexes.Add(lastColumn);
                    lastColumn++;
                }
            }
        }

        private IDictionary<string, int> GetRowKeys(ExcelWorksheet sheet)
        {
            var result = new Dictionary<string, int>();

            var uniqueRowIdBuilder = new List<string>();
            for (var rowIndex = _headerRowsCount + 1; rowIndex <= sheet.Dimension.Rows; rowIndex++)
            {
                var nameCellValue = sheet.Cells[rowIndex, WorksheetHelpers.RowNameColumnIndex].Value as string;
                if (WorksheetHelpers.IsGroupRow(sheet, rowIndex, _totalColumnIndexes) &&
                    !WorksheetHelpers.IsTotalRow(sheet, rowIndex, _leftPaneWidth + 1))
                {
                    uniqueRowIdBuilder.Add(nameCellValue);

                    var groupRowKey = string.Join("-", uniqueRowIdBuilder);
                    result.Add(groupRowKey, rowIndex);
                    continue;
                }

                if (WorksheetHelpers.IsTotalRow(sheet, rowIndex, _leftPaneWidth + 1))
                {
                    var rowKeyPrefix = string.Join("-", uniqueRowIdBuilder);
                    var uniqueMeasures = new Dictionary<string, int>();
                    for (var shift = 0; shift < _measuresCount; shift++)
                    {
                        var measureRowIndex = rowIndex + shift;
                        var measureCellValue = GetMeasureCellValue(sheet, measureRowIndex);
                        if (!uniqueMeasures.ContainsKey(measureCellValue))
                        {
                            uniqueMeasures.Add(measureCellValue, 0);
                        }

                        uniqueMeasures[measureCellValue]++;
                        var counter = uniqueMeasures[measureCellValue];
                        var uniqueMeasureValue = counter > 1 ? $"{measureCellValue} ({counter})" : measureCellValue;
                        var totalRowKey = $"{rowKeyPrefix}-{nameCellValue}-{uniqueMeasureValue}";

                        result.Add(totalRowKey, measureRowIndex);
                    }

                    if (uniqueRowIdBuilder.Any())
                    {
                        uniqueRowIdBuilder.RemoveAt(uniqueRowIdBuilder.Count - 1);
                    }

                    rowIndex += _measuresCount - 1;
                    continue;
                }

                if (!WorksheetHelpers.IsDataRow(sheet, rowIndex))
                {
                    continue;
                }

                var currentIdPrefix = string.Join("-", uniqueRowIdBuilder);
                var currentId = $"{currentIdPrefix}-{nameCellValue}";

                result.Add(currentId, rowIndex);
                rowIndex += _measuresCount - 1;
            }

            return result;
        }

        private void FormatSummaryRows(ExcelWorksheet sheet)
        {
            var firstDataRowIndex = _headerRowsCount + 1;
            var columnsCount = sheet.Dimension.Columns;

            var isPreviousRowTotal = false;
            for (var rowIndex = firstDataRowIndex; rowIndex <= sheet.Dimension.Rows; rowIndex++)
            {
                if (WorksheetHelpers.IsTotalRow(sheet, rowIndex, _leftPaneWidth + 1))
                {
                    isPreviousRowTotal = true;
                    continue;
                }

                var nameCell = sheet.Cells[rowIndex, WorksheetHelpers.RowNameColumnIndex];
                if (isPreviousRowTotal)
                {
                    if (nameCell.Value == null)
                    {
                        continue;
                    }

                    isPreviousRowTotal = false;
                }

                if (WorksheetHelpers.IsGroupRow(sheet, rowIndex, _totalColumnIndexes))
                {
                    continue;
                }

                var rowsToDelete = new List<int>();
                if (WorksheetHelpers.IsDataRow(sheet, rowIndex))
                {
                    for (var shift = 1; shift < _measuresCount; shift++)
                    {
                        var measureRowIndex = rowIndex + shift;
                        rowsToDelete.Add(measureRowIndex);

                        for (var x = _leftPaneWidth + 1; x <= columnsCount; x++)
                        {
                            var nextMeasureCell = sheet.Cells[measureRowIndex, x];
                            if (_totalColumnIndexes.Contains(x)
                                || WorksheetHelpers.IsEmptyCell(nextMeasureCell))
                            {
                                continue;
                            }

                            sheet.Cells[rowIndex, x].Value = nextMeasureCell.Value;
                        }
                    }
                }

                if (_isMeasureColumnExists)
                {
                    sheet.Cells[rowIndex, WorksheetHelpers.RowMeasureColumnIndex].Value = null;
                }

                foreach (var rowToDelete in rowsToDelete.OrderByDescending(x => x))
                {
                    sheet.DeleteRow(rowToDelete);
                }

                var color = _neutralColorGenerator.GetNextColor();

                for (var x = _leftPaneWidth + 1; x <= columnsCount; x++)
                {
                    var cell = sheet.Cells[rowIndex, x];
                    if (!_totalColumnIndexes.Contains(x)
                        && WorksheetHelpers.IsEmptyCell(cell))
                    {
                        cell.Value = null;
                        continue;
                    }

                    if (cell.Value == null || _totalColumnIndexes.Contains(x))
                    {
                        continue;
                    }

                    cell.Value = null;
                    cell.Style.Fill.BackgroundColor.SetColor(color);

                    if (!_cellsWithData.ContainsKey(rowIndex))
                    {
                        _cellsWithData[rowIndex] = new HashSet<int>();
                    }

                    _cellsWithData[rowIndex].Add(x);
                }
            }
        }

        private void ProcessGroups(ExcelWorksheet sheet,
                                   int startRowIndex,
                                   int endRowIndex,
                                   int outlineLevel)
        {
            if (outlineLevel > 0 && outlineLevel <= 7)
            {
                for (var i = startRowIndex; i <= endRowIndex; i++)
                {
                    sheet.Row(i).OutlineLevel = outlineLevel;
                    sheet.Row(i).Collapsed = true;
                }
            }

            for (var rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
            {
                if (!WorksheetHelpers.IsGroupRow(sheet, rowIndex, _totalColumnIndexes))
                {
                    continue;
                }

                var startGroupRowIndex = rowIndex;
                var groupName = sheet.Cells[rowIndex, WorksheetHelpers.RowNameColumnIndex].Value.ToString().Trim();

                var endGroupRowIndex = startGroupRowIndex;
                var isEndGroupFound = false;
                var duplicatesCount = 0;
                for (var i = startGroupRowIndex + 1; i <= sheet.Dimension.Rows; i++)
                {
                    if (isEndGroupFound)
                    {
                        endGroupRowIndex = i - 1;
                        break;
                    }

                    var nameCell = sheet.Cells[i, WorksheetHelpers.RowNameColumnIndex];
                    if (nameCell.Value == null)
                    {
                        continue;
                    }

                    if (nameCell.Value.ToString().Trim() == groupName && !WorksheetHelpers.IsDataRow(sheet, i))
                    {
                        duplicatesCount++;
                    }

                    if (!WorksheetHelpers.IsTotalRow(sheet, i, _leftPaneWidth + 1) ||
                        WorksheetHelpers.IsDataRow(sheet, i))
                    {
                        continue;
                    }

                    var totalName = sheet.Cells[i, WorksheetHelpers.RowNameColumnIndex].Value.ToString().Trim();
                    if (totalName != $"{WorksheetHelpers.TotalRowIndicator} {groupName}")
                    {
                        continue;
                    }

                    if (duplicatesCount == 0)
                    {
                        isEndGroupFound = true;
                    }
                    else
                    {
                        duplicatesCount--;
                    }
                }
                
                ProcessGroups(sheet, startGroupRowIndex + 1, endGroupRowIndex - 1, outlineLevel + 1);
                rowIndex = endGroupRowIndex;

                var rowsWithData = _cellsWithData.Keys.Where(x => x > startGroupRowIndex && x <= endGroupRowIndex)
                                                 .OrderBy(x => x);
                if (!rowsWithData.Any())
                {
                    continue;
                }

                var color = _neutralColorGenerator.GetNextColor();
                var columnsWithData = rowsWithData.SelectMany(x => _cellsWithData[x]).Distinct();
                foreach (var column in columnsWithData)
                {
                    var cell = sheet.Cells[startGroupRowIndex, column];
                    cell.Value = null;
                    cell.Style.Fill.BackgroundColor.SetColor(color);
                }
            }
        }
    }
}