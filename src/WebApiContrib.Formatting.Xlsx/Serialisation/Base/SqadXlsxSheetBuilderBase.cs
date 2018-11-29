using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base
{
    public abstract class SqadXlsxSheetBuilderBase
    {
        private readonly bool _shouldAutoFit;

        protected readonly List<DataTable> SheetTables;
        protected DataTable CurrentTable;

        public ExcelColumnInfoCollection SheetColumns { get; private set; }

        protected SqadXlsxSheetBuilderBase(string sheetName, bool isReferenceSheet = false, bool shouldAutoFit = true)
        {
            IsReferenceSheet = isReferenceSheet;
            _shouldAutoFit = shouldAutoFit;

            SheetTables = new List<DataTable>();

            CurrentTable = new DataTable(sheetName);
            SheetTables.Add(CurrentTable);

            SheetColumns = new ExcelColumnInfoCollection();
        }

        public bool IsReferenceSheet { get; }

        public void AddAndActivateNewTable(string sheetName)
        {
            CurrentTable = new DataTable(sheetName);
            SheetTables.Add(CurrentTable);
        }

        public void AppendColumns(DataColumnCollection columns)
        {
            foreach (DataColumn c in columns)
            {
                CurrentTable.Columns.Add(c.ColumnName, typeof(ExcelCell));
            }
        }

        public void AppendRow(IEnumerable<ExcelCell> row)
        {
            var dataRow = CurrentTable.NewRow();
            foreach (var cell in row)
            {
                dataRow.SetField(cell.CellHeader, cell);
            }

            CurrentTable.Rows.Add(dataRow);
        }

        public void AppendColumnHeaderRowItem(ExcelColumnInfo column)
        {
            SheetColumns.Add(column);
            var headerName = column.IsExcelHeaderDefined ? column.Header : column.PropertyName;
            var dc = new DataColumn(headerName, typeof(ExcelCell));

            if (column.IsHidden)
                dc.ColumnMapping = MappingType.Hidden;

            CurrentTable.Columns.Add(dc);
        }

        public bool ContainsTable(string sheetName)
        {
            return SheetTables.Any(x => x.TableName == sheetName);
        }

        public virtual void CompileSheet(ExcelPackage package)
        {
            ExcelWorksheet worksheet;
            if (IsReferenceSheet)
            {
                worksheet = package.Workbook.Worksheets.Add("Reference");
                worksheet.Hidden = eWorkSheetHidden.VeryHidden;
            }
            else
            {
                worksheet = package.Workbook.Worksheets.Add(CurrentTable.TableName);
            }

            PreCompileActions();

            foreach (var table in SheetTables)
            {
                CompileSheet(worksheet, table);

                if (!_shouldAutoFit)
                {
                    continue;
                }

                //worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                foreach (DataColumn col in table.Columns)
                {
                    if (worksheet.Name.Equals("Reference"))
                    {
                        break;
                    }

                    if (col.ColumnMapping == MappingType.Hidden)
                    {
                        worksheet.Column(col.Ordinal + 1).Hidden = true;
                    }
                }
            }

            PostCompileActions(worksheet);
        }

        protected virtual void PreCompileActions()
        {
        }

        protected virtual void PostCompileActions(ExcelWorksheet worksheet)
        {
        }

        protected abstract void CompileSheet(ExcelWorksheet worksheet, DataTable table);
    }
}