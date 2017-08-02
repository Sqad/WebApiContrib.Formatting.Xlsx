using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebApiContrib.Formatting.Xlsx.Interfaces;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class SqadXlsxSerialiser : IXlsxSerialiser
    {
        private IColumnResolver _columnResolver { get; set; }
        private ISheetResolver _sheetResolver { get; set; }
        private Func<string, DataTable> _staticValuesResolver { get; set; }
        public bool IgnoreFormatting => false;

        public SqadXlsxSerialiser(Func<string, DataTable> staticValuesResolver) : this(new DefaultSheetResolver(), new DefaultColumnResolver(), staticValuesResolver)
        {

        }

        public SqadXlsxSerialiser(ISheetResolver sheetResolver, IColumnResolver columnResolver, Func<string, DataTable> StaticValuesResolver)
        {
            _sheetResolver = sheetResolver;
            _columnResolver = columnResolver;
            this._staticValuesResolver = StaticValuesResolver;
        }

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return true;
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName = null)//, SqadXlsxSheetBuilder sheetBuilder)
        {
            var columnInfo = _columnResolver.GetExcelColumnInfo(itemType, value, sheetName);

            SqadXlsxSheetBuilder sheetBuilder = null;

            if (sheetName == null)
            {
                var sheetAttribute = itemType.GetCustomAttributes(true).SingleOrDefault(s => s is Attributes.ExcelSheetAttribute);
                sheetName = sheetAttribute != null ? (sheetAttribute as Attributes.ExcelSheetAttribute).SheetName : itemType.Name;
            }

            if (columnInfo.Count() > 0)
            {
                sheetBuilder = new SqadXlsxSheetBuilder(sheetName);
                sheetBuilder.AppendHeaderRow(columnInfo.Select(s => s.Header));
                document.AppendSheet(sheetBuilder);
            }

            //adding rows data
            if (value != null)
            {
                var columns = columnInfo.Keys.ToList();

                if (value is IEnumerable<object> && (value as IEnumerable<object>).Count() > 0)
                {
                    foreach (var dataObj in value as IEnumerable<object>)
                    {
                        PopulateRows(columns, dataObj, sheetBuilder, columnInfo, document);
                        //CheckColumnsForResolveSheets(document, columnInfo);
                        var deepSheetsInfo = _sheetResolver.GetExcelSheetInfo(itemType, dataObj);
                        PopulateInnerObjectSheets(deepSheetsInfo, document, itemType);
                    }
                }
                else if (!(value is IEnumerable<object>))
                {
                    PopulateRows(columns, value, sheetBuilder, columnInfo, document);
                    //CheckColumnsForResolveSheets(document, columnInfo);
                    var sheetsInfo = _sheetResolver.GetExcelSheetInfo(itemType, value);
                    PopulateInnerObjectSheets(sheetsInfo, document, itemType);
                }
            }

            if (sheetBuilder != null)
                sheetBuilder.ShouldAutoFit = true;
        }

        private void PopulateRows(List<string> columns, object value, SqadXlsxSheetBuilder sheetBuilder, ExcelColumnInfoCollection columnInfo = null, IXlsxDocumentBuilder document = null)
        {

            if (sheetBuilder == null)
                return;

            var row = new List<ExcelCell>();

            for (int i = 0; i <= columns.Count - 1; i++)
            {
                ExcelCell cell = new ExcelCell();

                string columnName = columns[i];
                object lookUpObject = value;
                if (columnName.Contains(":"))
                {
                    string[] columnPath = columnName.Split(':');
                    columnName = columnPath.Last();

                    for (int l = 1; l < columnPath.Count() - 1; l++)
                    {
                        lookUpObject = FormatterUtils.GetFieldOrPropertyValue(lookUpObject, columnPath[l]);
                    }
                }

                var cellValue = GetFieldOrPropertyValue(lookUpObject, columnName);

                ExcelColumnInfo info = null;
                if (columnInfo != null)
                {
                    info = columnInfo[i];
                    #region Reference Row
                    if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveFromTable) == false && _staticValuesResolver != null)
                    {
                        DataTable columntResolveTable = _staticValuesResolver(info.ExcelColumnAttribute.ResolveFromTable);
                        columntResolveTable.TableName = info.ExcelColumnAttribute.ResolveFromTable;
                        if (string.IsNullOrEmpty(info.ExcelColumnAttribute.OverrideResolveTableName) == false)
                            columntResolveTable.TableName = info.ExcelColumnAttribute.OverrideResolveTableName;

                        cell.DataValidationSheet = columntResolveTable.TableName;

                        var referenceSheet = document.GetReferenceSheet();

                        if (referenceSheet == null)
                        {
                            referenceSheet = new SqadXlsxSheetBuilder(cell.DataValidationSheet, true);
                            document.AppendSheet(referenceSheet);
                        }
                        else
                        {
                            referenceSheet.AddAndActivateNewTable(cell.DataValidationSheet);
                        }

                        cell.DataValidationBeginRow = referenceSheet.GetNextAvailalbleRow();

                        this.PopulateReferenceSheet(referenceSheet, columntResolveTable);

                        cell.DataValidationRowsCount = referenceSheet.GetCurrentRowCount;

                        if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveName) == false)
                            cell.DataValidationNameCellIndex = referenceSheet.GetColumnIndexByColumnName(info.ExcelColumnAttribute.ResolveName);

                        if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveValue) == false)
                            cell.DataValidationValueCellIndex = referenceSheet.GetColumnIndexByColumnName(info.ExcelColumnAttribute.ResolveValue);

                    }
                    #endregion Reference Row
                }

                cell.CellValue = FormatCellValue(cellValue, info);

                row.Add(cell);
            }
            sheetBuilder.AppendRow(row.ToList());
        }

        //private void CheckColumnsForResolveSheets(IXlsxDocumentBuilder document, ExcelColumnInfoCollection columnInfo)
        //{
        //    if (columnInfo == null)
        //        return;

        //    foreach (var cInfo in columnInfo)
        //    {
        //        if (string.IsNullOrEmpty(cInfo.ExcelColumnAttribute.ResolveFromTable) == false && _staticValuesResolver != null)
        //        {
        //            DataTable columntResolveTable = _staticValuesResolver(cInfo.ExcelColumnAttribute.ResolveFromTable);
        //            columntResolveTable.TableName = cInfo.ExcelColumnAttribute.ResolveFromTable;
        //            if (string.IsNullOrEmpty(cInfo.ExcelColumnAttribute.OverrideResolveTableName) == false)
        //                columntResolveTable.TableName = cInfo.ExcelColumnAttribute.OverrideResolveTableName;

        //            var referenceSheetBuilder = this.PopulateReferenceSheet(document,columntResolveTable);

        //            document.AppendSheet(referenceSheetBuilder);

        //            columntResolveTable = null;
        //        }
        //    }
        //}

        private void PopulateReferenceSheet(SqadXlsxSheetBuilder referenceSheet, DataTable ReferenceSheet)
        {
            List<string> sheetResolveColumns = new List<string>();

            foreach (DataColumn c in ReferenceSheet.Columns)
            {
                sheetResolveColumns.Add(c.Caption);
            }

            SqadXlsxSheetBuilder sb = new SqadXlsxSheetBuilder(ReferenceSheet.TableName, true);
            sb.AppendHeaderRow(sheetResolveColumns);
            sb.ShouldAutoFit = true;

            foreach (DataRow r in ReferenceSheet.Rows)
            {
                Dictionary<string, object> resolveRow = new Dictionary<string, object>();


                foreach (DataColumn c in ReferenceSheet.Columns)
                {

                    if (c.DataType == typeof(int))
                    {
                        resolveRow.Add(c.Caption, Convert.ToInt32(r[c]));
                    }
                    else if (c.DataType == typeof(DateTime))
                    {
                        resolveRow.Add(c.Caption, Convert.ToDateTime(r[c]));
                    }
                    else
                        resolveRow.Add(c.Caption, r[c].ToString());
                }

                this.PopulateRows(sheetResolveColumns, resolveRow, sb);
            }

            //return sb;
        }

        private void PopulateInnerObjectSheets(ExcelSheetInfoCollection sheetsInfo, IXlsxDocumentBuilder document, Type itemType)
        {
            foreach (var sheet in sheetsInfo)
            {
                if (!(sheet.ExcelSheetAttribute is Attributes.ExcelSheetAttribute))
                    continue;

                string sheetName = sheet.ExcelSheetAttribute != null ? (sheet.ExcelSheetAttribute as Attributes.ExcelSheetAttribute).SheetName : itemType.Name;
                if (sheetName == null)
                    sheetName = sheet.SheetName;

                //sheetBuilder = new SqadXlsxSheetBuilder(document.AppendSheet(sheetName));

                this.Serialise(sheet.SheetType, sheet.SheetObject, document, sheetName);//, sheetBuilder);

                //sheetBuilder = null;
            }
        }

        protected virtual object GetFieldOrPropertyValue(object rowObject, string name)
        {
            var rowValue = FormatterUtils.GetFieldOrPropertyValue(rowObject, name);

            if (rowValue is DateTimeOffset)
                return FormatterUtils.ConvertFromDateTimeOffset((DateTimeOffset)rowValue);

            else if (rowObject is Dictionary<string, object>)
                return (rowObject as Dictionary<string, object>)[name];

            else if (FormatterUtils.IsExcelSupportedType(rowValue))
                return rowValue;

            else if ((rowValue is IEnumerable<object>))
                return string.Join(",", rowValue as IEnumerable<object>);

            return rowValue == null || DBNull.Value.Equals(rowValue)
                ? string.Empty
                : rowValue.ToString();
        }

        protected virtual object FormatCellValue(object cellValue, ExcelColumnInfo info = null)
        {
            if (info != null)
            {
                // Boolean transformations.
                if (info.ExcelColumnAttribute != null && info.ExcelColumnAttribute.TrueValue != null && cellValue.Equals("True"))
                    return info.ExcelColumnAttribute.TrueValue;

                else if (info.ExcelColumnAttribute != null && info.ExcelColumnAttribute.FalseValue != null && cellValue.Equals("False"))
                    return info.ExcelColumnAttribute.FalseValue;

                else if (!string.IsNullOrWhiteSpace(info.FormatString) & string.IsNullOrEmpty(info.ExcelNumberFormat))
                    return string.Format(info.FormatString, cellValue);

                else if (cellValue.GetType() == typeof(DateTime))
                    return string.Format("{0:MM/dd/yyyy}", cellValue);
            }

            return cellValue;
        }
    }
}
