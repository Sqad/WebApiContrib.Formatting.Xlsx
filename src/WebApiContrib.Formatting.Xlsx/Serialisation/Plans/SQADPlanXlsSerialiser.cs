using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using SQAD.MTNext.Business.Models.Attributes;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.Services.Repositories.Export;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans
{
    public class SQADPlanXlsSerialiser : IXlsxSerialiser
    {
        public SerializerType SerializerType => SerializerType.Default;

        private IColumnResolver _columnResolver { get; set; }
        private ISheetResolver _sheetResolver { get; set; }
        private IExportHelpersRepository _staticValuesResolver { get; set; }
        public bool IgnoreFormatting => false;

        public SQADPlanXlsSerialiser(IExportHelpersRepository staticValuesResolver)
            : this(new DefaultSheetResolver(), new DefaultColumnResolver(), staticValuesResolver)
        {

        }

        public SQADPlanXlsSerialiser(ISheetResolver sheetResolver,
                                     IColumnResolver columnResolver,
                                     IExportHelpersRepository StaticValuesResolver)
        {
            _sheetResolver = sheetResolver;
            _columnResolver = columnResolver;
            this._staticValuesResolver = StaticValuesResolver;
        }

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            
            return valueType == typeof(ChartData);
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName = null)//, SqadXlsxSheetBuilder sheetBuilder)
        {
            ExcelColumnInfoCollection columnInfo = _columnResolver.GetExcelColumnInfo(itemType, value, sheetName);

            SqadXlsxPlanSheetBuilder sheetBuilder = null;

            if (sheetName == null)
            {
                var sheetAttribute = itemType.GetCustomAttributes(true).SingleOrDefault(s => s is ExcelSheetAttribute);
                sheetName = sheetAttribute != null ? (sheetAttribute as ExcelSheetAttribute).SheetName : itemType.Name;
            }

            if (columnInfo.Any())
            {
                sheetBuilder = document.GetSheetByName(sheetName) as SqadXlsxPlanSheetBuilder;

                if (sheetBuilder == null)
                {
                    sheetBuilder = new SqadXlsxPlanSheetBuilder(sheetName);

                    //Convert Dictionary Column
                    foreach (var col in columnInfo)
                    {
                        if (col.PropertyName.EndsWith("_Dict_"))
                        {
                            string columnName = col.PropertyName.Replace("_Dict_", "");

                            Dictionary<int, double> colValueDict = GetFieldOrPropertyValue(value, col.PropertyName.Replace("_Dict_", "")) as Dictionary<int, double>;
                            if (columnName.Contains(":") && (colValueDict == null || (colValueDict != null && string.IsNullOrEmpty(colValueDict.ToString()))))
                            {
                                colValueDict = GetFieldPathValue(value, columnName) as Dictionary<int, double>;
                            }

                            if (colValueDict == null)
                                continue;

                            int dictColumnCount = colValueDict.Count();

                            for (int i = 1; i <= dictColumnCount; i++)
                            {
                                ExcelColumnInfo temlKeyColumn = col.Clone() as ExcelColumnInfo;
                                temlKeyColumn.PropertyName = temlKeyColumn.PropertyName.Replace("_Dict_", $":Key:{i}");
                                sheetBuilder.AppendColumnHeaderRowItem(temlKeyColumn);

                                ExcelColumnInfo temlValueColumn = col.Clone() as ExcelColumnInfo;
                                temlValueColumn.PropertyName = temlValueColumn.PropertyName.Replace("_Dict_", $":Value:{i}");
                                sheetBuilder.AppendColumnHeaderRowItem(temlValueColumn);
                            }
                        }
                        else if (col.PropertyName.EndsWith("_CustomField_"))
                        {
                            string columnName = col.PropertyName.Replace("_CustomField_", "");

                            List<object> colCustomFields = GetFieldOrPropertyValue(value, col.PropertyName.Replace("_CustomField_", "")) as List<object>;
                            if (columnName.Contains(":") && (colCustomFields == null || (colCustomFields != null && string.IsNullOrEmpty(colCustomFields.ToString()))))
                            {
                                colCustomFields = GetFieldPathValue(value, columnName) as List<object>;
                            }

                            foreach (var customField in colCustomFields)
                            {
                                int customFieldId = ((dynamic)customField).ID;
                                ExcelColumnInfo temlKeyColumn = col.Clone() as ExcelColumnInfo;
                                temlKeyColumn.PropertyName = temlKeyColumn.PropertyName.Replace("_CustomField_", $":{customFieldId}");

                                string customFieldDef = _staticValuesResolver.GetCustomField(customFieldId);
                                temlKeyColumn.ExcelColumnAttribute.Header = temlKeyColumn.Header= temlKeyColumn.PropertyName + ":" + customFieldDef;

                                sheetBuilder.AppendColumnHeaderRowItem(temlKeyColumn);
                            }
                        }
                        else
                        {
                            sheetBuilder.AppendColumnHeaderRowItem(col);
                        }
                    }

                    //sheetBuilder.AppendColumns(columnInfo);

                    document.AppendSheet(sheetBuilder);
                }
            }

            if (sheetName != null && sheetBuilder == null)
            {
                sheetBuilder = (SqadXlsxPlanSheetBuilder)document.GetSheetByName(sheetName);
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
                        var deepSheetsInfo = _sheetResolver.GetExcelSheetInfo(itemType, dataObj);
                        PopulateInnerObjectSheets(deepSheetsInfo, document, itemType);
                    }
                }
                else if (!(value is IEnumerable<object>))
                {
                    PopulateRows(columns, value, sheetBuilder, columnInfo, document);
                    var sheetsInfo = _sheetResolver.GetExcelSheetInfo(itemType, value);
                    PopulateInnerObjectSheets(sheetsInfo, document, itemType);
                }
            }

            if (sheetBuilder != null)
            {
                sheetBuilder.ShouldAddHeaderRow = true;
            }
            
        }

        private void PopulateRows(List<string> columns, object value, SqadXlsxPlanSheetBuilder sheetBuilder, ExcelColumnInfoCollection columnInfo = null, IXlsxDocumentBuilder document = null)
        {

            if (sheetBuilder == null)
                return;

            var row = new List<ExcelCell>();

            for (int i = 0; i <= columns.Count - 1; i++)
            {
                string columnName = columns[i];

                if (columnName.EndsWith("_Dict_"))
                {
                    columnName = columnName.Replace("_Dict_", "");

                    Dictionary<int, double> dictValue = new Dictionary<int, double>();

                    if (columnName.Contains(":"))
                    {
                        var valueObject = GetFieldPathValue(value, columnName);
                        if (valueObject != null && string.IsNullOrEmpty(valueObject.ToString()) == false)
                        {
                            dictValue = (Dictionary<int, double>)valueObject;
                        }
                    }
                    else
                    {
                        dictValue = (Dictionary<int, double>)GetFieldOrPropertyValue(value, columnName);
                    }

                    int colCount = 1;
                    foreach (var kv in dictValue)
                    {
                        ExcelCell keyCell = new ExcelCell();
                        keyCell.CellHeader = columnName + $":Key:{colCount}";
                        keyCell.CellValue = kv.Key;
                        ExcelColumnInfo info = null;
                        if (columnInfo != null)
                        {
                            info = columnInfo[i];
                            CreateReferenceCell(info, columnName, document, ref keyCell);
                        }
                        row.Add(keyCell);



                        ExcelCell valueCell = new ExcelCell();
                        valueCell.CellHeader = columnName + $":Value:{colCount}";
                        valueCell.CellValue = kv.Value;
                        row.Add(valueCell);

                        colCount++;
                    }
                }
                else if (columnName.EndsWith("_CustomField_"))
                {
                    columnName = columnName.Replace("_CustomField_", "");

                    List<object> customFields = null;
                    if (columnName.Contains(":"))
                    {
                        var valueObject = GetFieldPathValue(value, columnName);
                        if (valueObject != null && string.IsNullOrEmpty(valueObject.ToString()) == false)
                        {
                            customFields = (List<object>)valueObject;
                        }
                    }
                    else
                    {
                        customFields = (List<object>)GetFieldOrPropertyValue(value, columnName);
                    }

                    foreach(var customField in customFields)
                    {
                        string columnNameCombined = columnName + ":" + ((dynamic)customField).ID;

                        var customFieldColumnInfo = sheetBuilder.SheetColumns.Where(w => w.PropertyName == columnNameCombined).FirstOrDefault();

                        ExcelCell customValueHeaderCell = new ExcelCell();
                        customValueHeaderCell.CellHeader = customFieldColumnInfo.Header;
                        customValueHeaderCell.CellValue = ((dynamic)customField).Value;
                        row.Add(customValueHeaderCell);
                    }

                }
                else
                {
                    ExcelCell cell = new ExcelCell();

                    cell.CellHeader = columnName;

                    var cellValue = GetFieldOrPropertyValue(value, columnName);

                    if (columnName.Contains(":") && (cellValue == null || (cellValue != null && string.IsNullOrEmpty(cellValue.ToString()))))
                    {
                        cellValue = GetFieldPathValue(value, columnName);
                    }

                    ExcelColumnInfo info = null;
                    if (columnInfo != null)
                    {
                        info = columnInfo[i];

                        CreateReferenceCell(info, columnName, document, ref cell);
                    }

                    if (cellValue != null)
                        cell.CellValue = FormatCellValue(cellValue, info);

                    if (info != null)
                    {
                        if (info.IsExcelHeaderDefined)
                            cell.CellHeader = info.Header;
                    }

                    row.Add(cell);
                }
            }

            if (row.Count() > 0)
                sheetBuilder.AppendRow(row.ToList());
        }

        private void PopulateReferenceSheet(SqadXlsxPlanSheetBuilder referenceSheet, DataTable ReferenceSheet)
        {
            //SqadXlsxSheetBuilder sb = new SqadXlsxSheetBuilder(ReferenceSheet.TableName, true);
            referenceSheet.AppendColumns(ReferenceSheet.Columns);


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
                        if(r[c] is System.DBNull)
                            resolveRow.Add(c.Caption, string.Empty);
                        else
                           resolveRow.Add(c.Caption, Convert.ToDateTime(r[c]));
                    }
                    else
                        resolveRow.Add(c.Caption, r[c].ToString());
                }

                PopulateRows(resolveRow.Keys.ToList(), resolveRow, referenceSheet);
            }

            //return sb;
        }

        private void PopulateInnerObjectSheets(ExcelSheetInfoCollection sheetsInfo, IXlsxDocumentBuilder document, Type itemType)
        {
            foreach (var sheet in sheetsInfo)
            {
                if (!(sheet.ExcelSheetAttribute is ExcelSheetAttribute))
                    continue;

                string sheetName = sheet.ExcelSheetAttribute != null ? (sheet.ExcelSheetAttribute as ExcelSheetAttribute).SheetName : itemType.Name;
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

            List<object> itemsToProcess = null;
            if (rowObject is IEnumerable<object> && (rowObject as IEnumerable<object>).Count() > 0)
            {
                itemsToProcess = (rowObject as IEnumerable<object>).ToList();
            }
            else
            {
                itemsToProcess = new List<object>() { rowObject };
            }


            bool isResultList = false;
            bool isResultDictionary = false;
            bool isResultCustomField = false;

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

                    if (result != null)
                    {
                        //if (index == pathSplit.Count())
                        //{
                        if (index != pathSplit.Count())
                        {
                            if (result.GetType().Name.StartsWith("List"))
                                itemsToProcess.AddRange(result as IEnumerable<object>);
                            else
                                itemsToProcess.Add(result);

                            continue;
                        }

                        if (result.GetType().Name.StartsWith("List"))
                        {

                            if (result.GetType().FullName.Contains("CustomFieldModel"))
                            {
                                List<int> ids = resultsList.Select(s => (int)((dynamic)s).ID).ToList();
                                var filtereCustomFields = (result as IEnumerable<object>).Where(w => ids.Contains(((dynamic)w).ID) == false).ToList();
                                resultsList.AddRange(filtereCustomFields);
                                isResultCustomField = true;
                            }
                            else
                            {
                                resultsList.AddRange(result as IEnumerable<object>);
                                isResultList = true;
                            }
                        }
                        else if (result.GetType().Name.StartsWith("Dictionary"))
                        {
                            resultsList.Add(result);
                            isResultDictionary = true;
                        }
                        else
                            resultsList.Add(result);

                        //}
                        //else
                        //{
                        //    if (result.GetType().Name.StartsWith("List"))
                        //        itemsToProcess.AddRange(result as IEnumerable<object>);
                        //    else
                        //        itemsToProcess.Add(result);
                        //}
                    }
                }
            }

            string returnString = string.Empty;


            if (isResultCustomField)
            {
                return resultsList;
            }
            else if (isResultList)
            {
                for (int i = 1; i <= resultsList.Count(); i++)
                {
                    returnString += $"{i}->{resultsList[i - 1]}, ";
                }
            }
            else if (isResultDictionary)
            {
                return resultsList.FirstOrDefault();
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

                else if (cellValue.GetType() == typeof(DateTime) || DateTime.TryParse(cellValue.ToString(), out var test))
                    return string.Format("{0:MM/dd/yyyy}", DateTime.Parse( cellValue.ToString()));

            }

            return cellValue;
        }

        public void CreateReferenceCell(ExcelColumnInfo info, string columnName, IXlsxDocumentBuilder document, ref ExcelCell cell)
        {
            if (_staticValuesResolver == null)
                return;

            DataTable columntResolveTable = null;


            if (info.PropertyType != null && info.PropertyType.BaseType == typeof(Enum))
            {
                columntResolveTable = _staticValuesResolver.GetRecordsFromEnum(info.PropertyType);
                info.ExcelColumnAttribute.ResolveFromTable = columnName;
            }
            else if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveFromTable) == false)
            {
                columntResolveTable = _staticValuesResolver.GetRecordsByTableName(info.ExcelColumnAttribute.ResolveFromTable); ;
            }

            if (columntResolveTable != null)
            {
                columntResolveTable.TableName = info.ExcelColumnAttribute.ResolveFromTable;
                if (string.IsNullOrEmpty(info.ExcelColumnAttribute.OverrideResolveTableName) == false)
                    columntResolveTable.TableName = info.ExcelColumnAttribute.OverrideResolveTableName;

                cell.DataValidationSheet = columntResolveTable.TableName;

                var referenceSheet = document.GetReferenceSheet() as SqadXlsxPlanSheetBuilder;

                if (referenceSheet == null)
                {
                    referenceSheet = new SqadXlsxPlanSheetBuilder(cell.DataValidationSheet, true);
                    document.AppendSheet(referenceSheet);
                }
                else
                {
                    referenceSheet.AddAndActivateNewTable(cell.DataValidationSheet);
                }

                cell.DataValidationBeginRow = referenceSheet.GetNextAvailableRow();

                this.PopulateReferenceSheet(referenceSheet, columntResolveTable);

                cell.DataValidationRowsCount = referenceSheet.GetCurrentRowCount;

                if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveName) == false)
                    cell.DataValidationNameCellIndex = referenceSheet.GetColumnIndexByColumnName(info.ExcelColumnAttribute.ResolveName);

                if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveValue) == false)
                    cell.DataValidationValueCellIndex = referenceSheet.GetColumnIndexByColumnName(info.ExcelColumnAttribute.ResolveValue);
            }
            else if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveFromTable) == false)
            {
                columntResolveTable = _staticValuesResolver.GetRecordsByTableName(info.ExcelColumnAttribute.ResolveFromTable); ;
            }

            if (columntResolveTable != null)
            {
                columntResolveTable.TableName = info.ExcelColumnAttribute.ResolveFromTable;
                if (string.IsNullOrEmpty(info.ExcelColumnAttribute.OverrideResolveTableName) == false)
                    columntResolveTable.TableName = info.ExcelColumnAttribute.OverrideResolveTableName;

                cell.DataValidationSheet = columntResolveTable.TableName;

                var referenceSheet = document.GetReferenceSheet() as SqadXlsxPlanSheetBuilder;

                if (referenceSheet == null)
                {
                    referenceSheet = new SqadXlsxPlanSheetBuilder(cell.DataValidationSheet, true);
                    document.AppendSheet(referenceSheet);
                }
                else
                {
                    referenceSheet.AddAndActivateNewTable(cell.DataValidationSheet);
                }

                cell.DataValidationBeginRow = referenceSheet.GetNextAvailableRow();

                this.PopulateReferenceSheet(referenceSheet, columntResolveTable);

                cell.DataValidationRowsCount = referenceSheet.GetCurrentRowCount;

                if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveName) == false)
                    cell.DataValidationNameCellIndex = referenceSheet.GetColumnIndexByColumnName(info.ExcelColumnAttribute.ResolveName);

                if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveValue) == false)
                    cell.DataValidationValueCellIndex = referenceSheet.GetColumnIndexByColumnName(info.ExcelColumnAttribute.ResolveValue);
            }
        }
    }
}
