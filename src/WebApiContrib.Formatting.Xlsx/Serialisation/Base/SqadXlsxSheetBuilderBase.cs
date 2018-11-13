using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base
{
    public abstract class SqadXlsxSheetBuilderBase
    {
        private readonly bool _shouldAutoFit;

        protected readonly List<DataTable> SheetTables;
        protected DataTable CurrentTable;

        protected SqadXlsxSheetBuilderBase(string sheetName, bool isReferenceSheet = false, bool shouldAutoFit = true)
        {
            IsReferenceSheet = isReferenceSheet;
            _shouldAutoFit = shouldAutoFit;

            SheetTables = new List<DataTable>();

            CurrentTable = new DataTable(sheetName);
            SheetTables.Add(CurrentTable);
        }

        public bool IsReferenceSheet { get; }

        public void AddAndActivateNewTable(string sheetName)
        {
            CurrentTable = new DataTable(sheetName);
            SheetTables.Add(CurrentTable);
        }

        public void AppendColumns(ExcelColumnInfoCollection columns)
        {
            foreach (var col in columns)
            {
                var headerName = col.IsExcelHeaderDefined ? col.Header : col.PropertyName;
                var dataColumn = new DataColumn(headerName, typeof(ExcelCell));

                if (col.IsHidden)
                {
                    dataColumn.ColumnMapping = MappingType.Hidden;
                }

                CurrentTable.Columns.Add(dataColumn);
            }
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

            PreCompileActions(worksheet);

            foreach (var table in SheetTables)
            {
                CompileSheet(worksheet, table);

                if (!_shouldAutoFit)
                {
                    continue;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    
                foreach (DataColumn col in table.Columns)
                {
                    if (worksheet.Name.Equals("Reference")) break;

                    if (col.ColumnMapping == MappingType.Hidden)
                    {
                        worksheet.Column(col.Ordinal + 1).Hidden = true;
                    }
                }
            }

            PostCompileActions(worksheet);
        }

        protected virtual void PreCompileActions(ExcelWorksheet worksheet)
        {

        }

        protected virtual void PostCompileActions(ExcelWorksheet worksheet)
        {

        }

        protected abstract void CompileSheet(ExcelWorksheet worksheet, DataTable table);
    }
}
