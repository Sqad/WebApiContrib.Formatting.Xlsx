using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class SqadXlsxSerialiser : IXlsxSerialiser
    {
        private IColumnResolver _columnResolver { get; set; }
        private ISheetResolver _sheetResolver { get; set; }

        public bool IgnoreFormatting => false;

        public SqadXlsxSerialiser() : this(new DefaultSheetResolver(), new DefaultColumnResolver()) { }

        public SqadXlsxSerialiser(ISheetResolver sheetResolver, IColumnResolver columnResolver)
        {
            _sheetResolver = sheetResolver;
            _columnResolver = columnResolver;
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
                sheetBuilder = new SqadXlsxSheetBuilder(document.AppendSheet(sheetName));
                sheetBuilder.AppendRow(columnInfo.Select(s => s.Header));
            }

            //adding rows data
            if (value != null)
            {
                var columns = columnInfo.Keys.ToList();

                if (value is IEnumerable<object> && (value as IEnumerable<object>).Count() > 0)
                {
                    foreach (var dataObj in value as IEnumerable<object>)
                    {
                        PopulateRows(columns, dataObj, sheetBuilder, columnInfo);
                        var deepSheetsInfo = _sheetResolver.GetExcelSheetInfo(itemType, dataObj);
                        PopulateInnerObjectSheets(deepSheetsInfo, document, itemType, sheetBuilder);
                    }
                }
                else if (!(value is IEnumerable<object>))
                {
                    PopulateRows(columns, value, sheetBuilder, columnInfo);
                    var sheetsInfo = _sheetResolver.GetExcelSheetInfo(itemType, value);
                    PopulateInnerObjectSheets(sheetsInfo, document, itemType, sheetBuilder);
                }
            }

            if (sheetBuilder != null)
                sheetBuilder.AutoFit();
        }

        private void PopulateRows(List<string> columns, object value, SqadXlsxSheetBuilder sheetBuilder, ExcelColumnInfoCollection columnInfo)
        {
            if (sheetBuilder == null)
                return;

            var row = new List<object>();

            for (int i = 0; i <= columns.Count - 1; i++)
            {
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
                var info = columnInfo[i];
                row.Add(FormatCellValue(cellValue, info));

            }

            sheetBuilder.AppendRow(row.ToList());
        }

        private void PopulateInnerObjectSheets(ExcelSheetInfoCollection sheetsInfo, IXlsxDocumentBuilder document, Type itemType, SqadXlsxSheetBuilder sheetBuilder)
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

            else if (FormatterUtils.IsExcelSupportedType(rowValue))
                return rowValue;

            return rowValue == null || DBNull.Value.Equals(rowValue)
                ? string.Empty
                : rowValue.ToString();
        }

        protected virtual object FormatCellValue(object cellValue, ExcelColumnInfo info)
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

            else
                return cellValue;
        }
    }
}
