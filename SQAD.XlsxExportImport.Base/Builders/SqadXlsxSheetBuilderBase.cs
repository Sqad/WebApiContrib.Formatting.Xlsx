using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeOpenXml;
using SQAD.XlsxExportImport.Base.Models;

namespace SQAD.XlsxExportImport.Base.Builders
{
    public abstract class SqadXlsxSheetBuilderBase
    {
        private readonly bool _shouldAutoFit;

        protected readonly List<DataTable> SheetTables;
        protected DataTable CurrentTable;

        public ExcelColumnInfoCollection SheetColumns { get; private set; }

        protected SqadXlsxSheetBuilderBase(string sheetName, bool isReferenceSheet = false, bool isPreservationSheet = false, bool isHidden = false, bool shouldAutoFit = true)
        {
            IsReferenceSheet = isReferenceSheet;
            IsPreservationSheet = isPreservationSheet;
            IsHidden = isHidden;

            _shouldAutoFit = shouldAutoFit;

            SheetTables = new List<DataTable>();

            CurrentTable = new DataTable(sheetName);
            SheetTables.Add(CurrentTable);

            SheetColumns = new ExcelColumnInfoCollection();
        }

        public bool IsReferenceSheet { get; }

        public bool IsPreservationSheet { get; }

        public bool IsHidden { get; set; }

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

        public void AppendRow(DataRowCollection rows)
        {
            foreach (DataRow row in rows)
            {
                var dataRow = CurrentTable.NewRow();
                for (var i = 0; i < CurrentTable.Columns.Count; i++)
                {
                    dataRow.SetField(CurrentTable.Columns[i].ColumnName, row[i]);
                }
                CurrentTable.Rows.Add(dataRow);
            }
        }

        public void AppendColumnHeaderRowItem(string columnName)
        {
            var dc = new DataColumn(columnName, typeof(ExcelCell));
            CurrentTable.Columns.Add(dc);
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
                if (IsHidden)
                    worksheet.Hidden = eWorkSheetHidden.VeryHidden;
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