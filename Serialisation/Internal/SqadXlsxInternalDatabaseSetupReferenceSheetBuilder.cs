using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OfficeOpenXml;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Parsing.Attributes;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Parsing.Base;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Result;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Internal
{
    internal class SqadXlsxInternalDatabaseSetupReferenceSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private const int HeaderRowsCount = 3;
        private readonly IEnumerable<ExportResultItem<ExcelRowBase>> _exportResults;

        private static readonly MethodInfo WriteCellMethodInfo;

        static SqadXlsxInternalDatabaseSetupReferenceSheetBuilder()
        {
            WriteCellMethodInfo = typeof(SqadXlsxInternalDatabaseSetupReferenceSheetBuilder)
                .GetMethod(nameof(WriteCell), BindingFlags.NonPublic | BindingFlags.Static);
        }

        public SqadXlsxInternalDatabaseSetupReferenceSheetBuilder(IEnumerable<ExportResultItem<ExcelRowBase>> exportResults)
            : base(string.Empty)
        {
            _exportResults = exportResults;
        }

        public override void CompileSheet(ExcelPackage package)
        {
            var workbook = package.Workbook;

            foreach (var exportResult in _exportResults)
            {
                var type = exportResult.RowType;
                if (!(type.GetCustomAttributes(typeof(ExcelParseSheetAttribute), false)
                          .FirstOrDefault() is ExcelParseSheetAttribute sheetAttribute))
                {
                    throw new ArgumentException($"Type '{type.FullName}' does not declared " +
                                                $"with '{nameof(ExcelParseSheetAttribute)}' attribute.");
                }

                if (string.IsNullOrWhiteSpace(sheetAttribute.SheetName))
                {
                    throw new
                        ArgumentException($"Type '{type.FullName}' has no specified " +
                                          $"'{nameof(sheetAttribute.SheetName)}' for " +
                                          $"'{nameof(ExcelParseSheetAttribute)}' attribute.");
                }

                var sheet = workbook.Worksheets[sheetAttribute.SheetName];
                FillSheet(sheet, type, exportResult);

                //note vv: force GC to free memory since each IEnumerable can be huge
                GC.Collect();
            }
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            throw new NotImplementedException();
        }

        private static void FillSheet(ExcelWorksheet sheet, Type rowType, ExportResultItem<ExcelRowBase> result)
        {
            var (fillRowActions, columns) = BuildFillRowAction(rowType);
            var columnsLookup = BuildColumnsLookup(sheet, columns.Keys);

            var currentRow = HeaderRowsCount + 1;
            try
            {
                foreach (var row in result.Rows)
                {
                    foreach (var fillRowAction in fillRowActions)
                    {
                        fillRowAction(sheet, row, currentRow, columnsLookup);
                    }

                    currentRow++;
                }
            }
            catch (Exception e)
            {
                sheet.TabColor = Color.Red;
                var errorCell = sheet.Cells[1, sheet.Dimension.Columns + 1];

                errorCell.Value = e.Message;
                errorCell.Style.WrapText = false;

                return;
            }
            

            if (currentRow == HeaderRowsCount + 1)
            {
                return;
            }

            foreach (var (key, value) in columns)
            {
                if (value != typeof(DateTime))
                {
                    continue;
                }

                var columnIndex = columnsLookup[key];
                var cells = sheet.Cells[HeaderRowsCount + 1, columnIndex, sheet.Dimension.Rows, columnIndex];
                cells.Style.Numberformat.Format = "mm-dd-yy";
            }
        }

        private static (
            ICollection<Action<ExcelWorksheet, ExcelRowBase, int, IDictionary<string, int>>>,
            IDictionary<string, Type>
            ) BuildFillRowAction(Type type)
        {
            var sheetParameter = Expression.Parameter(typeof(ExcelWorksheet), "sheet");
            var rowIndexParameter = Expression.Parameter(typeof(int), "rowIndex");
            var columnsLookupParameter = Expression.Parameter(typeof(IDictionary<string, int>), "columnsLookup");

            var rowParameter = Expression.Parameter(typeof(ExcelRowBase), "row");

            var properties = type.GetProperties();
            var columns = new Dictionary<string, Type>();
            var propertyExpressions = new List<Expression>();
            foreach (var propertyInfo in properties)
            {
                if (!(propertyInfo.GetCustomAttributes(typeof(ExcelExportColumnAttribute), false)
                                  .FirstOrDefault() is ExcelExportColumnAttribute columnAttribute))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(columnAttribute.ColumnName))
                {
                    throw new ArgumentException($"Property '{propertyInfo.Name}' from type '{type.FullName}' " +
                                                $"has no specified '{nameof(columnAttribute.ColumnName)}' for " +
                                                $"'{nameof(ExcelExportColumnAttribute)}' attribute.");
                }

                if (!propertyInfo.GetMethod.IsPublic)
                {
                    throw new ArgumentException($"Property '{propertyInfo.Name}' from type " +
                                                $"'{type.FullName}' has no public getter.");
                }

                var valueParameter = Expression.Convert(Expression.Property(Expression.Convert(rowParameter, type),
                                                                            propertyInfo),
                                                        typeof(object));
                var columnNameParameter = Expression.Constant(columnAttribute.ColumnName);

                var writeInstruction = Expression.Call(WriteCellMethodInfo,
                                                       sheetParameter,
                                                       rowIndexParameter,
                                                       columnNameParameter,
                                                       columnsLookupParameter,
                                                       valueParameter);
                propertyExpressions.Add(writeInstruction);

                columns.Add(columnAttribute.ColumnName, propertyInfo.PropertyType);
            }

            var lambdas =
                propertyExpressions.Select(x => Expression
                                               .Lambda<Action<ExcelWorksheet, ExcelRowBase, int,
                                                   IDictionary<string, int>>>(x,
                                                                              sheetParameter,
                                                                              rowParameter,
                                                                              rowIndexParameter,
                                                                              columnsLookupParameter))
                                   .Select(x=>x.Compile())
                                   .ToList();

            return (lambdas, columns);
        }

        private static Dictionary<string, int> BuildColumnsLookup(ExcelWorksheet sheet, ICollection<string> columnNames)
        {
            var columnsLookup = new Dictionary<string, int>(columnNames.Count);
            for (var columnIndex = 1; columnIndex <= sheet.Dimension.Columns; columnIndex++)
            {
                var cell = sheet.Cells[HeaderRowsCount - 1, columnIndex];
                var cellValue = cell.Value as string;

                if (string.IsNullOrWhiteSpace(cellValue)
                    || !columnNames.Contains(cellValue))
                {
                    continue;
                }

                columnsLookup.Add(cellValue, columnIndex);
            }

            return columnsLookup;
        }

        private static void WriteCell(ExcelWorksheet sheet,
                                      int rowIndex,
                                      string columnName,
                                      IDictionary<string, int> columnsLookup,
                                      object value)
        {
            if (!columnsLookup.TryGetValue(columnName, out var columnIndex))
            {
                return;
            }

            sheet.Cells[rowIndex, columnIndex].Value = value;
        }
    }
}