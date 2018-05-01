﻿using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Resources = SQAD.MTNext.Resources.Properties.Resources;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx
{
    public class SqadXlsxSheetBuilder
    {
        private readonly int __ROWS_BETWEEN_REFERENCE_SHEETS__ = 2;
        private DataTable _currentTable { get; set; }

        public string CurrentTableName => _currentTable.TableName;

        private bool _isReferenceSheet { get; set; }
        public bool IsReferenceSheet => _isReferenceSheet;

        public bool ShouldAutoFit { get; set; }

        public bool ShouldAddHeaderRow { get; set; }

        private List<DataTable> _sheetTables { get; set; }

        public List<DataTable> SheetTables => _sheetTables;

        public SqadXlsxSheetBuilder(string sheetName, bool isReferenceSheet = false)
        {
            _isReferenceSheet = isReferenceSheet;
            _sheetTables = new List<DataTable>();

            _currentTable = new DataTable(sheetName);
            _sheetTables.Add(_currentTable);
        }

        public void AddAndActivateNewTable(string sheetName)
        {
            _currentTable = new DataTable(sheetName);
            _sheetTables.Add(_currentTable);
        }

        public void AppendColumnHeaderRow(ExcelColumnInfoCollection columns)
        {
            _currentTable.Columns.AddRange(columns.Select(s => new DataColumn(s.PropertyName, typeof(ExcelCell))).ToArray());
        }

        public void AppendColumnHeaderRow(DataColumnCollection columns)
        {
            foreach (DataColumn c in columns)
            {
                _currentTable.Columns.Add(c.ColumnName, typeof(ExcelCell));
            }
        }

        public void AppendRow(IEnumerable<ExcelCell> row)
        {
            DataRow dRow = _currentTable.NewRow();
            foreach (var cell in row)
            {
                //if (string.IsNullOrEmpty(cell.CellValue.ToString()) == false)
                dRow.SetField(cell.CellHeader, cell);
            }
            _currentTable.Rows.Add(dRow);
        }

        public int GetNextAvailalbleRow()
        {
            //to get next available row, we get total of all rows plus number of all reference sheets 
            //multiply by two (rows between the sheets)
            int totalRows = _sheetTables.Select(s => s.Rows.Count).Sum();
            if (totalRows == 0)
                return 1 + __ROWS_BETWEEN_REFERENCE_SHEETS__; // completely new reference sheet no need to shift from top
            else
                //total rows  plus number or tables adjustment for emptry row increast everytime plus number of sheets multiply rows between sheets
                return totalRows + _sheetTables.Count() + (_sheetTables.Count() * __ROWS_BETWEEN_REFERENCE_SHEETS__);//already rows there, need to make space for next reference table
        }

        public int GetCurrentRowCount => _currentTable.Rows.Count;

        public int GetColumnIndexByColumnName(string columnName)
        {
            int index = 0;

            var column = _currentTable.Columns[columnName];
            if (column != null)
                index = column.Ordinal + 1;

            return index;
        }

        public void CompileSheet(ExcelPackage package)
        {
            if (_sheetTables.Count() == 0)
                return;

            List<string> sheetCodeColumnStatements = new List<string>();
            ExcelWorksheet worksheet = null;

            if (_isReferenceSheet == true)
            {
                worksheet = package.Workbook.Worksheets.Add("Reference");
                worksheet.Hidden = eWorkSheetHidden.VeryHidden;
            }
            else
                worksheet = package.Workbook.Worksheets.Add(_currentTable.TableName);

            int rowCount = ShouldAddHeaderRow ? 3 : 1;

            foreach (var table in _sheetTables)
            {
                if (_isReferenceSheet == true)
                {
                    var mergeTitleCell = worksheet.Cells[rowCount, 1, rowCount, 3];
                    mergeTitleCell.Value = table.TableName;
                    mergeTitleCell.Merge = true;
                    mergeTitleCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    mergeTitleCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                    rowCount++;
                }

                if (ShouldAddHeaderRow == true)
                {
                    var headerRow = worksheet.Row(1);
                    headerRow.Height = 45.60;

                    #region First Cell Logo
                    worksheet.Column(1).Width = 10.22;
                    var cell = worksheet.Cells[1, 1];
                    var picture = worksheet.Drawings.AddPicture("SQADLogo", Resources.Properties.Resources.SQADLogo);
                    picture.SetPosition(cell.Rows - 1, 2, cell.Columns - 1, 15);
                    #endregion First Cell Logo

                    #region Second Tab Name
                    var cells = worksheet.Cells[1, 2, 1, 6];
                    cells.Merge = true;

                    var tabName = cells.RichText.Add($"{CurrentTableName.ToUpper()} ");
                    var staticNameTabText = cells.RichText.Add("DATA FIELDS");

                    tabName.Size = staticNameTabText.Size = 40;
                    tabName.Color = staticNameTabText.Color = System.Drawing.Color.FromArgb(0, 159, 220);
                    tabName.FontName = staticNameTabText.FontName = "Calibri";
                    tabName.Bold = true;

                    #endregion Second Tab Name

                }

                foreach (DataColumn col in table.Columns)
                    worksheet.Cells[rowCount, col.Ordinal + 1].Value = col.ColumnName;

                foreach (DataRow row in table.Rows)
                {
                    rowCount++;
                    foreach (DataColumn col in table.Columns)
                    {
                        var colObject = row[col];
                        int excelColumnIndex = col.Ordinal + 1; //adjustment for excel column count index (start from 1. nondevelopment count, duh!)
                        if (!(colObject is ExcelCell))
                        {
                            worksheet.Cells[rowCount, excelColumnIndex].Value = colObject;
                        }
                        else if (colObject is ExcelCell)
                        {
                            ExcelCell cell = colObject as ExcelCell;

                            worksheet.Cells[rowCount, excelColumnIndex].Value = cell.CellValue;

                            if (!string.IsNullOrEmpty(cell.DataValidationSheet))
                            {
                                var dataValidation = worksheet.DataValidations.AddListValidation(worksheet.Cells[rowCount, excelColumnIndex].Address);
                                dataValidation.ShowErrorMessage = true;

                                string validationAddress = $"'Reference'!{worksheet.Cells[cell.DataValidationBeginRow, cell.DataValidationNameCellIndex, (cell.DataValidationBeginRow + cell.DataValidationRowsCount) - 1, cell.DataValidationNameCellIndex]}";
                                dataValidation.Formula.ExcelFormula = validationAddress;

                                string code = string.Empty;
                                code += $"If Target.Column = {excelColumnIndex} Then \n";
                                code += $"   matchVal = Application.Match(Target.Value, Worksheets(\"Reference\").Range(\"{worksheet.Cells[cell.DataValidationBeginRow, cell.DataValidationNameCellIndex, (cell.DataValidationBeginRow + cell.DataValidationRowsCount) - 1, cell.DataValidationNameCellIndex].Address}\"), 0) \n";
                                code += $"   selectedNum = Application.Index(Worksheets(\"Reference\").Range(\"{worksheet.Cells[cell.DataValidationBeginRow, cell.DataValidationValueCellIndex, (cell.DataValidationBeginRow + cell.DataValidationRowsCount) - 1, cell.DataValidationValueCellIndex].Address}\"), matchVal, 1) \n";
                                code += "   If Not IsError(selectedNum) Then \n";
                                code += "       Target.Value = selectedNum \n";
                                code += "   End If \n";
                                code += "End If \n";

                                sheetCodeColumnStatements.Add(code);
                            }
                            else if (cell.CellValue != null && bool.TryParse(cell.CellValue.ToString(), out var result))
                            {
                                var dataValidation = worksheet.DataValidations.AddListValidation(worksheet.Cells[rowCount, excelColumnIndex].Address);
                                dataValidation.ShowErrorMessage = true;
                                dataValidation.Formula.Values.Add("True");
                                dataValidation.Formula.Values.Add("False");
                            }
                        }
                    }
                }

                rowCount += __ROWS_BETWEEN_REFERENCE_SHEETS__;

            }





            if (worksheet.Dimension != null && ShouldAutoFit)
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            #region sheet code to resolve reference column
            if (sheetCodeColumnStatements.Count() > 0)
            {
                string worksheetOnChangeCode = string.Empty;
                worksheetOnChangeCode += "Private Sub Worksheet_Change(ByVal Target As Range) \n";
                worksheetOnChangeCode += "  If Target.Value = Empty Then \n";
                worksheetOnChangeCode += "      Exit Sub \n";
                worksheetOnChangeCode += "  End If \n";

                foreach (var codePiece in sheetCodeColumnStatements)
                {
                    worksheetOnChangeCode += codePiece;
                }

                worksheetOnChangeCode += "End Sub";

                if (worksheet.Workbook.VbaProject == null)
                    worksheet.Workbook.CreateVBAProject();
                worksheet.CodeModule.Code = worksheetOnChangeCode;
            }
            #endregion sheet code to resolve reference column

        }
    }
}
