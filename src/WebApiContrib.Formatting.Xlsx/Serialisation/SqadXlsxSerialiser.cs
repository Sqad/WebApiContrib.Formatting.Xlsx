﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
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
            ExcelColumnInfoCollection columnInfo = _columnResolver.GetExcelColumnInfo(itemType, value, sheetName);

            SqadXlsxSheetBuilder sheetBuilder = null;

            if (sheetName == null)
            {
                var sheetAttribute = itemType.GetCustomAttributes(true).SingleOrDefault(s => s is SQAD.MTNext.Business.Models.Attributes.ExcelSheetAttribute);
                sheetName = sheetAttribute != null ? (sheetAttribute as SQAD.MTNext.Business.Models.Attributes.ExcelSheetAttribute).SheetName : itemType.Name;
            }

            if (columnInfo.Count() > 0)
            {
                sheetBuilder = document.GetSheetByName(sheetName);

                if (sheetBuilder == null)
                {
                    sheetBuilder = new SqadXlsxSheetBuilder(sheetName);
                    sheetBuilder.AppendHeaderRow(columnInfo);
                    document.AppendSheet(sheetBuilder);
                }
            }

            if (sheetName != null && sheetBuilder == null)
            {
                sheetBuilder = document.GetSheetByName(sheetName);
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

                cell.CellHeader = columnName;

                //bool lookUpObjectIsList = false;
                //object lookUpObject = value;
                //if (columnName.Contains(":"))
                //{
                //    string[] columnPath = columnName.Split(':');
                //    columnName = columnPath.Last();

                //    for (int l = 1; l < columnPath.Count() - 1; l++)
                //    {
                //        lookUpObject = FormatterUtils.GetFieldOrPropertyValue(lookUpObject, columnPath[l]);

                //        if (lookUpObject != null && lookUpObject.GetType().Name.StartsWith("List"))
                //        {
                //            lookUpObjectIsList = true;
                //            break;
                //        }

                //    }
                //}

                //if (lookUpObjectIsList)
                //{
                //    this.Serialise(FormatterUtils.GetEnumerableItemType(lookUpObject.GetType()), lookUpObject as IEnumerable<object>, document, sheetBuilder.CurrentTableName);

                //}
                //else
                //{
                var cellValue = GetFieldOrPropertyValue(value, columnName);

                if (columnName.Contains(":") && (cellValue == null || (cellValue != null && string.IsNullOrEmpty(cellValue.ToString()))))
                {
                    cellValue = GetFieldPathValue(value, columnName);
                }

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

                if (cellValue != null)
                    cell.CellValue = FormatCellValue(cellValue, info);

                row.Add(cell);
                //}
            }

            if (row.Count() > 0)
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
            //SqadXlsxSheetBuilder sb = new SqadXlsxSheetBuilder(ReferenceSheet.TableName, true);
            referenceSheet.AppendHeaderRow(ReferenceSheet.Columns);
            referenceSheet.ShouldAutoFit = true;


            foreach (DataRow r in ReferenceSheet.Rows)
            {
                Dictionary<string, object> resolveRow = new Dictionary<string, object>();

                foreach (DataColumn c in ReferenceSheet.Columns)
                {
                    if (c.DataType == typeof(int))
                    {
                        int i = 0;
                        if (int.TryParse(r[c].ToString(), out i))
                            resolveRow.Add(c.Caption, Convert.ToInt32(i));
                    }
                    else if (c.DataType == typeof(DateTime))
                    {
                        resolveRow.Add(c.Caption, Convert.ToDateTime(r[c]));
                    }
                    else
                        resolveRow.Add(c.Caption, r[c].ToString());
                }

                this.PopulateRows(resolveRow.Keys.ToList(), resolveRow, referenceSheet);
            }

            //return sb;
        }

        private void PopulateInnerObjectSheets(ExcelSheetInfoCollection sheetsInfo, IXlsxDocumentBuilder document, Type itemType)
        {
            foreach (var sheet in sheetsInfo)
            {
                if (!(sheet.ExcelSheetAttribute is SQAD.MTNext.Business.Models.Attributes.ExcelSheetAttribute))
                    continue;

                string sheetName = sheet.ExcelSheetAttribute != null ? (sheet.ExcelSheetAttribute as SQAD.MTNext.Business.Models.Attributes.ExcelSheetAttribute).SheetName : itemType.Name;
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
                return rowValue;
            // return string.Join(",", rowValue as IEnumerable<object>);

            return rowValue == null || DBNull.Value.Equals(rowValue)
                ? string.Empty
                : rowValue.ToString();
        }

        private object GetFieldPathValue(object rowObject, string path)
        {
            string[] pathSplit = path.Split(':').Skip(1).ToArray();

            List<object> resultsList = new List<object>();

            List<object> itemsToProcess = new List<object>() { rowObject };

            bool isResultList = false;

            int index = 0;
            foreach (string objName in pathSplit)
            {
                index++;

                var tempItemsToProcess = itemsToProcess.ToList();
                itemsToProcess.Clear();

                foreach (var obj in tempItemsToProcess)
                {
                    if (obj == null)
                        continue;

                    Type type = obj.GetType();

                    System.Reflection.MemberInfo member = null;

                    var matchingProperties = type.GetProperties().Where(w => w.Name == objName).ToList();
                    if (matchingProperties.Count() > 0)
                    {
                        member = matchingProperties.First();
                    }
                    else if ((matchingProperties = type.GetProperties().Where(w => w.Name == "Value").ToList()).Count() > 1)
                    {
                        //property overwriten, and must take first
                        member = matchingProperties.First();
                    }
                    else
                    {
                        member = type.GetField(objName) ?? type.GetProperty(objName) as System.Reflection.MemberInfo;
                    }
                    if (member == null)
                    {
                        return null;
                    }
                    var result = new object();

                    switch (member.MemberType)
                    {
                        case MemberTypes.Property:
                            result = ((PropertyInfo)member).GetValue(obj, null);
                            break;
                        case MemberTypes.Field:
                            result = ((FieldInfo)member).GetValue(obj);
                            break;
                        default:
                            result = null;
                            break;
                    }

                    if (index == pathSplit.Count())
                    {
                        if (result != null)
                        {
                            if (result.GetType().Name.StartsWith("List"))
                            {
                                List<object> list = new List<object>();
                                var enumerator = ((System.Collections.IEnumerable)result).GetEnumerator();
                                while (enumerator.MoveNext())
                                {
                                    resultsList.Add(enumerator.Current);
                                }
                                isResultList = true;
                            }
                            else
                                resultsList.Add(result);
                        }
                    }
                    else
                    {
                        if (result != null)
                        {
                            if (result.GetType().Name.StartsWith("List"))
                                itemsToProcess.AddRange(result as IEnumerable<object>);
                            else
                                itemsToProcess.Add(result);
                        }
                    }
                }
            }

            string returnString = string.Empty;

            if (isResultList)
            {
                for (int i = 1; i <= resultsList.Count(); i++)
                {
                    returnString += $"{i}->{resultsList[i - 1]}, ";
                }
            }
            else
            {
                returnString = resultsList.Count() > 0 ? resultsList.First().ToString() : string.Empty;
            }
            return returnString;
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
