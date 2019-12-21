using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OfficeOpenXml;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Parsing;
using SQAD.MTNext.Business.Models.Internal.DatabaseSetup.Parsing.Attributes;
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
            : base(string.Empty, true)
        {
            _exportResults = exportResults;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            var workbook = worksheet.Workbook;

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

                GC.Collect();
            }
        }

        private static void FillSheet(ExcelWorksheet sheet, Type rowType, ExportResultItem<ExcelRowBase> result)
        {
            var (fillRowActions, columns) = BuildFillRowAction(rowType);
            var columnsLookup = BuildColumnsLookup(sheet, columns);

            var currentRow = HeaderRowsCount + 1;
            foreach (var row in result.Rows)
            {
                foreach (var fillRowAction in fillRowActions)
                {
                    fillRowAction(sheet, row, currentRow, columnsLookup);
                }

                currentRow++;
            }
        }

        private static (
            IEnumerable<Action<ExcelWorksheet, ExcelRowBase, int, IDictionary<string, int>>>,
            HashSet<string>
            ) BuildFillRowAction(Type type)
        {
            var sheetParameter = Expression.Parameter(typeof(ExcelWorksheet), "sheet");
            var rowIndexParameter = Expression.Parameter(typeof(int), "rowIndex");
            var columnsLookupParameter = Expression.Parameter(typeof(IDictionary<string, int>), "columnsLookup");

            var rowParameter = Expression.Parameter(typeof(ExcelRowBase), "row");

            var properties = type.GetProperties();
            var columns = new HashSet<string>();
            var propertyExpressions = new List<Expression>();
            foreach (var propertyInfo in properties)
            {
                if (!(propertyInfo.GetCustomAttributes(typeof(ExcelParseColumnAttribute), false)
                                  .FirstOrDefault() is ExcelParseColumnAttribute columnAttribute))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(columnAttribute.ColumnName))
                {
                    throw new ArgumentException($"Property '{propertyInfo.Name}' from type '{type.FullName}' " +
                                                $"has no specified '{nameof(columnAttribute.ColumnName)}' for " +
                                                $"'{nameof(ExcelParseColumnAttribute)}' attribute.");
                }

                if (!propertyInfo.GetMethod.IsPublic)
                {
                    throw new ArgumentException($"Property '{propertyInfo.Name}' from type " +
                                                $"'{type.FullName}' has no public getter.");
                }

                var valueParameter = Expression.Property(Expression.Convert(rowParameter, type), propertyInfo);
                var columnNameParameter = Expression.Constant(columnAttribute.ColumnName);

                var writeInstruction = Expression.Call(WriteCellMethodInfo,
                                                       sheetParameter,
                                                       rowIndexParameter,
                                                       columnNameParameter,
                                                       columnsLookupParameter,
                                                       valueParameter);
                propertyExpressions.Add(writeInstruction);

                columns.Add(columnAttribute.ColumnName);
            }

            var lambdas =
                propertyExpressions.Select(x => Expression
                                               .Lambda<Action<ExcelWorksheet, ExcelRowBase, int,
                                                   IDictionary<string, int>>>(x,
                                                                              sheetParameter,
                                                                              rowParameter,
                                                                              rowIndexParameter,
                                                                              columnsLookupParameter));

            return (lambdas.Select(x => x.Compile()), columns);
        }

        private static Dictionary<string, int> BuildColumnsLookup(ExcelWorksheet sheet, ICollection<string> columnNames)
        {
            var columnsLookup = new Dictionary<string, int>(columnNames.Count);
            for (var columnIndex = 1; columnIndex <= sheet.Dimension.Columns; columnIndex++)
            {
                var cell = sheet.Cells[HeaderRowsCount - 1, columnIndex];
                var cellValue = cell.Value as string;

                if (!columnNames.Contains(cellValue)
                    || string.IsNullOrWhiteSpace(cellValue))
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