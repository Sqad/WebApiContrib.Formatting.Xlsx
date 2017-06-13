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


        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, SqadXlsxSheetBuilder sheetBuilder)
        {
            var columnInfo = _columnResolver.GetExcelColumnInfo(itemType, value);
            string sheetName = string.Empty;

            if (sheetBuilder == null)
            {
                var sheetAttribute = itemType.GetCustomAttributes(true).SingleOrDefault(s => s is Attributes.ExcelSheetAttribute);
                sheetName = sheetAttribute != null ? (sheetAttribute as Attributes.ExcelSheetAttribute).SheetName : itemType.Name;
                sheetBuilder = new SqadXlsxSheetBuilder(document.AppendSheet(sheetName));
            }

            if (columnInfo.Count() > 0)
                sheetBuilder.AppendRow(columnInfo.Select(s => s.Header));

            //adding rows data
            if (value != null)
            {
                var columns = columnInfo.Keys.ToList();

                if (value is IEnumerable<object>)
                {
                    foreach (var dataObj in value as IEnumerable<object>)
                    {
                        PopulateRows(columns, dataObj, sheetBuilder, columnInfo);
                        //this.Serialise(itemType, dataObj, document, sheetBuilder);
                        var deepSheetsInfo = _sheetResolver.GetExcelSheetInfo(itemType, dataObj);
                        PopulateInnerObjectSheets(deepSheetsInfo, document, itemType, sheetBuilder);
                    }
                }
                else
                {
                    PopulateRows(columns, value, sheetBuilder, columnInfo);
                    var sheetsInfo = _sheetResolver.GetExcelSheetInfo(itemType, value);
                    PopulateInnerObjectSheets(sheetsInfo, document, itemType, sheetBuilder);
                }
            }

            sheetBuilder.AutoFit();
        }

        private void PopulateRows(List<string> columns, object value, SqadXlsxSheetBuilder sheetBuilder, ExcelColumnInfoCollection columnInfo)
        {

            var row = new List<object>();

            for (int i = 0; i <= columns.Count - 1; i++)
            {
                var cellValue = GetFieldOrPropertyValue(value, columns[i]);
                var info = columnInfo[i];

                //row.Add(FormatCellValue(cellValue, info));
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

                sheetBuilder = new SqadXlsxSheetBuilder(document.AppendSheet(sheetName));

                this.Serialise(sheet.SheetType, sheet.SheetObject, document, sheetBuilder);

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

            else
                return cellValue;
        }
    }
}
