using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans
{
    public class SqadXlsxPlanSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly int __ROWS_BETWEEN_REFERENCE_SHEETS__ = 2;

        private readonly List<string> _sheetCodeColumnStatements;
        private int _rowsCount;

        

        public bool ShouldAddHeaderRow { private get; set; }

        public SqadXlsxPlanSheetBuilder(string sheetName, bool isReferenceSheet = false)
            : base(sheetName, isReferenceSheet)
        {
            _sheetCodeColumnStatements = new List<string>();
        }

        public int GetNextAvailableRow()
        {
            //to get next available row, we get total of all rows plus number of all reference sheets 
            //multiply by two (rows between the sheets)
            var totalRows = SheetTables.Select(s => s.Rows.Count).Sum();
            if (totalRows == 0)
            {
                // completely new reference sheet no need to shift from top
                return 1 + __ROWS_BETWEEN_REFERENCE_SHEETS__;
            }

            //already rows there, need to make space for next reference table
            return totalRows + SheetTables.Count + SheetTables.Count * __ROWS_BETWEEN_REFERENCE_SHEETS__;
        }

        public int GetCurrentRowCount => CurrentTable.Rows.Count;

        public int GetColumnIndexByColumnName(string columnName)
        {
            var index = 0;

            var column = CurrentTable.Columns[columnName];
            if (column != null)
            {
                index = column.Ordinal + 1;
            }

            return index;
        }

        protected override void PreCompileActions()
        {
            _sheetCodeColumnStatements.Clear();
            _rowsCount = ShouldAddHeaderRow ? 3 : 1;
        }

        protected override void PostCompileActions(ExcelWorksheet worksheet)
        {
            worksheet.Cells[3, 1, 3, worksheet.Dimension.Columns].Style.Fill.PatternType =
                OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[3, 1, 3, worksheet.Dimension.Columns].Style.Fill.BackgroundColor
                     .SetColor(System.Drawing.Color.FromArgb(242, 242, 242));

            worksheet.Cells[3, 1, 3, worksheet.Dimension.Columns].Style.Border.Top.Style =
                OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[3, 1, 3, worksheet.Dimension.Columns].Style.Border.Bottom.Style =
                OfficeOpenXml.Style.ExcelBorderStyle.Thick;
            worksheet.Cells[3, 1, 3, worksheet.Dimension.Columns].Style.Border.Bottom.Color
                     .SetColor(System.Drawing.Color.Black);

            if (!_sheetCodeColumnStatements.Any())
            {
                return;
            }

            var worksheetOnChangeCode = string.Empty;
            worksheetOnChangeCode += "Private Sub Worksheet_Change(ByVal Target As Range) \n";
            worksheetOnChangeCode += "  If Target.Value = Empty Then \n";
            worksheetOnChangeCode += "      Exit Sub \n";
            worksheetOnChangeCode += "  End If \n";

            foreach (var codePiece in _sheetCodeColumnStatements)
            {
                worksheetOnChangeCode += codePiece;
            }

            worksheetOnChangeCode += "End Sub";

            if (worksheet.Workbook.VbaProject == null)
            {
                worksheet.Workbook.CreateVBAProject();
            }

            worksheet.CodeModule.Code = worksheetOnChangeCode;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (IsReferenceSheet)
            {
                var mergeTitleCell = worksheet.Cells[_rowsCount, 1, _rowsCount, 3];
                mergeTitleCell.Value = table.TableName;
                mergeTitleCell.Merge = true;
                mergeTitleCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                mergeTitleCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                _rowsCount++;
            }

            if (ShouldAddHeaderRow)
            {
                worksheet.DefaultColWidth = 13;
                var headerRow = worksheet.Row(1);
                headerRow.Height = 45.60;

                #region Logo and Tab name

                var logoAndTabTitleCells = worksheet.Cells[1, 1, 1, 6];
                logoAndTabTitleCells.Merge = true;

                //Picture
                var picture = worksheet.Drawings.AddPicture("SQADLogo", Resources.Properties.Resources.SQADLogo);
                picture.SetPosition(0, 2, 0, 5);

                //Tab name text
                var tabName = logoAndTabTitleCells.RichText.Add($"{CurrentTable.TableName.ToUpper()} ");
                var staticNameTabText = logoAndTabTitleCells.RichText.Add("DATA FIELDS");

                tabName.Size = staticNameTabText.Size = 40;
                tabName.Color = staticNameTabText.Color = System.Drawing.Color.FromArgb(0, 159, 220);
                tabName.FontName = staticNameTabText.FontName = "Calibri";
                tabName.Bold = true;

                logoAndTabTitleCells.Style.Indent = 7;

                #endregion Logo and Tab name

                #region Notice text

                var noticeTextCells = worksheet.Cells[1, 7, 1, 8];
                noticeTextCells.Merge = true;

                var noticeImportantText = noticeTextCells.RichText.Add("IMPORTANT: ");
                noticeImportantText.Bold = true;
                noticeImportantText.Size = 11;
                noticeImportantText.FontName = "Calibri";
                noticeImportantText.Color = System.Drawing.Color.White;

                var noticeWarningText =
                    noticeTextCells
                        .RichText
                        .Add("Text instructions for how to complete this section of the page. (the first two lines of this sheet will be omitted from the import).");
                noticeWarningText.Size = 11;
                noticeWarningText.FontName = "Calibri";
                noticeWarningText.Bold = false;
                noticeWarningText.Color = System.Drawing.Color.White;

                noticeTextCells.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                noticeTextCells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(132, 151, 176));
                noticeTextCells.Style.WrapText = true;
                noticeTextCells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

                #endregion Notice text

                var fullLogoCells = worksheet.Cells[1, 9, 1, 12];
                fullLogoCells.Merge = true;

                var fullLogo =
                    worksheet.Drawings.AddPicture("SQADLogoFull", Resources.Properties.Resources.SQADLogoFull);
                fullLogo.SetPosition(0, 2, 9, 0);
            }


            foreach (DataColumn col in table.Columns)
            {
                if (worksheet.Name.Equals("Reference")) break;

                var colName = worksheet.Cells[_rowsCount, col.Ordinal + 1].RichText.Add(col.ColumnName);
                colName.Bold = true;
                colName.Size = 13;

                worksheet.Cells[_rowsCount, col.Ordinal + 1].Style.Border.Right.Style =
                    OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                worksheet.Cells[_rowsCount, col.Ordinal + 1].Style.Border.Right.Color
                         .SetColor(System.Drawing.Color.Black);

                if ((col.Ordinal + 1) % 2 == 0)
                {
                    int maxRows = _rowsCount + table.Rows.Count;

                    worksheet.Cells[_rowsCount, col.Ordinal + 1, maxRows, col.Ordinal + 1].Style.Fill.PatternType =
                        OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[_rowsCount, col.Ordinal + 1, maxRows, col.Ordinal + 1].Style.Fill.BackgroundColor
                             .SetColor(System.Drawing.Color.FromArgb(242, 242, 242));

                    worksheet.Cells[_rowsCount, col.Ordinal + 1, maxRows, col.Ordinal + 1].Style.Border.Right.Style =
                        OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[_rowsCount, col.Ordinal + 1, maxRows, col.Ordinal + 1].Style.Border.Left.Style =
                        OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells[_rowsCount, col.Ordinal + 1, maxRows, col.Ordinal + 1].Style.Border.Right.Color
                             .SetColor(System.Drawing.Color.Black);
                    worksheet.Cells[_rowsCount, col.Ordinal + 1, maxRows, col.Ordinal + 1].Style.Border.Left.Color
                             .SetColor(System.Drawing.Color.Black);
                }

                if (col.ColumnMapping == MappingType.Hidden)
                {
                    worksheet.Column(col.Ordinal + 1).Hidden = true;
                }
            }

            foreach (DataRow row in table.Rows)
            {
                _rowsCount++;

                foreach (DataColumn col in table.Columns)
                {
                    var colObject = row[col];

                    //adjustment for excel column count index (start from 1. nondevelopment count, duh!)
                    var excelColumnIndex = col.Ordinal + 1;
                    if (!(colObject is ExcelCell))
                    {
                        worksheet.Cells[_rowsCount, excelColumnIndex].Value = colObject;
                    }
                    else
                    {
                        var cell = colObject as ExcelCell;

                        worksheet.Cells[_rowsCount, excelColumnIndex].Value = cell.CellValue;


                        if (!string.IsNullOrEmpty(cell.DataValidationSheet))
                        {
                            var dataValidation =
                                worksheet.DataValidations.AddListValidation(worksheet
                                                                            .Cells[_rowsCount, excelColumnIndex]
                                                                            .Address);
                            dataValidation.ShowErrorMessage = true;

                            string validationAddress =
                                $"'Reference'!{worksheet.Cells[cell.DataValidationBeginRow, cell.DataValidationNameCellIndex, (cell.DataValidationBeginRow + cell.DataValidationRowsCount) - 1, cell.DataValidationNameCellIndex]}";
                            dataValidation.Formula.ExcelFormula = validationAddress;

                            string code = string.Empty;
                            code += $"If Target.Column = {excelColumnIndex} Then \n";
                            code +=
                                $"   matchVal = Application.Match(Target.Value, Worksheets(\"Reference\").Range(\"{worksheet.Cells[cell.DataValidationBeginRow, cell.DataValidationNameCellIndex, (cell.DataValidationBeginRow + cell.DataValidationRowsCount) - 1, cell.DataValidationNameCellIndex].Address}\"), 0) \n";
                            code +=
                                $"   selectedNum = Application.Index(Worksheets(\"Reference\").Range(\"{worksheet.Cells[cell.DataValidationBeginRow, cell.DataValidationValueCellIndex, (cell.DataValidationBeginRow + cell.DataValidationRowsCount) - 1, cell.DataValidationValueCellIndex].Address}\"), matchVal, 1) \n";
                            code += "   If Not IsError(selectedNum) Then \n";
                            code += "       Target.Value = selectedNum \n";
                            code += "   End If \n";
                            code += "End If \n";

                            _sheetCodeColumnStatements.Add(code);
                        }
                        else if (cell.CellValue != null && bool.TryParse(cell.CellValue.ToString(), out var result))
                        {
                            var dataValidation =
                                worksheet.DataValidations.AddListValidation(worksheet
                                                                            .Cells[_rowsCount, excelColumnIndex]
                                                                            .Address);
                            dataValidation.ShowErrorMessage = true;
                            dataValidation.Formula.Values.Add("True");
                            dataValidation.Formula.Values.Add("False");
                        }
                    }
                }
            }

            _rowsCount += __ROWS_BETWEEN_REFERENCE_SHEETS__;
        }
    }
}