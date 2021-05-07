using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using SQAD.XlsxExportImport.Base.Models;
using SQAD.XlsxExportImport.Base.Interfaces;
using SQAD.XlsxExportImport.Base.Serialization;
using SQAD.XlsxExportImport.Base.Attributes;
using SQAD.XlsxExportImport.Base.Formatters;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using SQAD.MTNext.Services.Repositories.Export;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;


namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans
{
    public class SQADPlanXlsSerializer : IXlsxSerializer
    {
        public SerializerType SerializerType => SerializerType.Default;

        private IColumnResolver _columnResolver { get; set; }
        private ISheetResolver _sheetResolver { get; set; }
        private IExportHelpersRepository _staticValuesResolver { get; set; }

        private bool _isExportJsonToXls;

        private Dictionary<string, DataTable> _resolveTables = new Dictionary<string, DataTable>();
        public bool IgnoreFormatting => false;

        public SQADPlanXlsSerializer(IExportHelpersRepository staticValuesResolver, IModelMetadataProvider modelMetadataProvider
               , bool isExportJsonToXls = false)
            : this(new DefaultSheetResolver(), new DefaultColumnResolver(modelMetadataProvider), staticValuesResolver, isExportJsonToXls)
        {

        }

        public SQADPlanXlsSerializer(ISheetResolver sheetResolver,
                                     IColumnResolver columnResolver,
                                     IExportHelpersRepository StaticValuesResolver,
                                     bool isExportJsonToXls = false)
        {
            _sheetResolver = sheetResolver;
            _columnResolver = columnResolver;
            this._staticValuesResolver = StaticValuesResolver;
            _isExportJsonToXls = isExportJsonToXls;
        }

        public bool CanSerializeType(Type valueType, Type itemType)
        {

            return valueType == typeof(ChartData);
        }

        public void Serialize(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName = null, string columnPrefix = null, XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder sheetBuilderOverride = null)
        {
            ExcelColumnInfoCollection columnInfo = _columnResolver.GetExcelColumnInfo(itemType, value, sheetName);

            XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder sheetBuilder = null;

            if (sheetName == null)
            {
                var sheetAttribute = itemType.GetCustomAttributes(true).SingleOrDefault(s => s is ExcelSheetAttribute);
                sheetName = sheetAttribute != null ? (sheetAttribute as ExcelSheetAttribute).SheetName : itemType.Name;
            }

            if (columnInfo.Any())
            {
                if (sheetBuilderOverride == null)
                    sheetBuilder = document.GetSheetByName(sheetName) as XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder;
                else
                    sheetBuilder = sheetBuilderOverride;

                if (sheetBuilder == null)
                {
                    sheetBuilder = new XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder(sheetName);
                    //Move this to attribute hidden property
                    //if (new List<string>() { "Formulas", "LeftTableColumn", "Cells" }.Contains(sheetName))
                    //{
                    //    sheetBuilder.IsHidden = true;
                    //}

                    document.AppendSheet(sheetBuilder);
                }

                //Convert Dictionary Column
                foreach (var col in columnInfo)
                {
                    if (col.PropertyName.EndsWith("_Dict_"))
                    {
                        string columnName = col.PropertyName.Replace("_Dict_", "");

                        object colValueDict = null;
                        if (sheetName == col.PropertyName.Replace("_Dict_", ""))
                        {
                            colValueDict = value;
                        } else
                        {
                            colValueDict = GetFieldOrPropertyValue(value, col.PropertyName.Replace("_Dict_", ""));
                        }

                        if (columnName.Contains(":") && (colValueDict == null || (colValueDict != null && string.IsNullOrEmpty(colValueDict.ToString()))))
                        {
                            colValueDict = GetFieldPathValue(value, columnName);
                        }

                        if (colValueDict == null || string.IsNullOrEmpty(colValueDict.ToString()))
                            continue;


                        object dictionaryKeys = colValueDict.GetType().GetProperty("Keys").GetValue(colValueDict);

                        int count = 0;
                        foreach (var key in (System.Collections.IEnumerable)dictionaryKeys)
                        {
                            ExcelColumnInfo temlKeyColumn = col.Clone() as ExcelColumnInfo;
                            temlKeyColumn.PropertyName = temlKeyColumn.PropertyName.Replace("_Dict_", $":Key:{count}");
                            sheetBuilder.AppendColumnHeaderRowItem(temlKeyColumn);

                            var currentItem = colValueDict.GetType().GetProperty("Item").GetValue(colValueDict, new object[] { key });

                            if (FormatterUtils.IsSimpleType(currentItem.GetType()))
                            {
                                ExcelColumnInfo temlValueColumn = col.Clone() as ExcelColumnInfo;
                                temlValueColumn.PropertyName = temlValueColumn.PropertyName.Replace("_Dict_", $":Value:{count}");
                                sheetBuilder.AppendColumnHeaderRowItem(temlValueColumn);
                            } else
                            {
                                string path = col.PropertyName.Replace("_Dict_", $":Value:{count}");
                                this.Serialize(currentItem.GetType(), value, document, sheetName, path, sheetBuilderOverride);
                            }

                            count++;
                        }
                    } else if (col.PropertyName.EndsWith("_List_"))
                    {
                        string columnName = col.PropertyName.Replace("_List_", "");

                        List<object> colListValue = GetFieldOrPropertyValue(value, col.PropertyName.Replace("_List_", "")) as List<object>;
                        if (columnName.Contains(":") && (colListValue == null || (colListValue != null && string.IsNullOrEmpty(colListValue.ToString()))))
                        {
                            colListValue = GetFieldPathValue(value, columnName) as List<object>;
                        }

                        if (colListValue == null)
                            continue;

                        int dictColumnCount = colListValue.Count();

                        for (int i = 0; i < dictColumnCount; i++)
                        {
                            string listColumnPrefix = col.PropertyName.Replace("_List_", $":{i}");
                            if (FormatterUtils.IsSimpleType(colListValue[i].GetType()))
                            {
                                ExcelColumnInfo colToAppend = (ExcelColumnInfo)col.Clone();
                                colToAppend.PropertyName = listColumnPrefix;
                                sheetBuilder.AppendColumnHeaderRowItem(colToAppend);
                            } else
                            {
                                this.Serialize(colListValue[i].GetType(), colListValue[i], document, null, listColumnPrefix, sheetBuilder);
                            }

                        }
                    } else if (col.PropertyName.EndsWith("_CustomField_") || col.PropertyName.EndsWith("_CustomField_Single_"))
                    {
                        string columnName = col.PropertyName.Replace("_CustomField_", "").Replace("Single_", "");

                        List<object> colCustomFields = GetFieldOrPropertyValue(value, columnName) as List<object>;
                        if (columnName.Contains(":") && (colCustomFields == null || (colCustomFields != null && string.IsNullOrEmpty(colCustomFields.ToString()))))
                        {
                            colCustomFields = GetFieldPathValue(value, columnName) as List<object>;
                        }

                        if (colCustomFields == null)
                            continue;

                        foreach (var customField in colCustomFields)
                        {

                            int customFieldId = ((dynamic)customField).ID;
                            bool isActual = ((dynamic)customField).Actual;

                            ExcelColumnInfo temlKeyColumn = col.Clone() as ExcelColumnInfo;

                            string propetyActual = isActual ? ":Actual" : string.Empty;

                            string customFieldDef = _staticValuesResolver.GetCustomFieldName(customFieldId);

                            if (col.PropertyName.EndsWith("_CustomField_Single_"))
                            {
                                customFieldDef = string.Empty;
                                temlKeyColumn.PropertyName = temlKeyColumn.PropertyName.Replace("_CustomField_Single_", $"{propetyActual}");
                                temlKeyColumn.ExcelColumnAttribute.Header = temlKeyColumn.Header = temlKeyColumn.PropertyName;

                            } else
                            {
                                temlKeyColumn.PropertyName = temlKeyColumn.PropertyName.Replace("_CustomField_", $"{propetyActual}:{customFieldId}");
                                temlKeyColumn.ExcelColumnAttribute.Header = temlKeyColumn.Header = temlKeyColumn.PropertyName + ":" + customFieldDef;
                            }




                            sheetBuilder.AppendColumnHeaderRowItem(temlKeyColumn);
                        }

                    } else
                    {
                        if (columnPrefix != null)
                        {
                            ExcelColumnInfo temlKeyColumn = col.Clone() as ExcelColumnInfo;
                            temlKeyColumn.PropertyName = $"{columnPrefix}:{temlKeyColumn.PropertyName}";
                            sheetBuilder.AppendColumnHeaderRowItem(temlKeyColumn);
                        } else
                        {
                            sheetBuilder.AppendColumnHeaderRowItem(col);
                        }
                    }
                }

            }

            //if its recursive do not populate rows and return to parent
            if (columnPrefix != null)
                return;


            if (sheetName != null && sheetBuilder == null)
            {
                sheetBuilder = (XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder)document.GetSheetByName(sheetName);
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
                } else if (!(value is IEnumerable<object>))
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

        private void PopulateRows(List<string> columns, object value, XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder sheetBuilder, ExcelColumnInfoCollection columnInfo = null, IXlsxDocumentBuilder document = null, List<ExcelCell> rowOverride = null)
        {

            if (sheetBuilder == null)
                return;

            List<ExcelCell> row = new List<ExcelCell>();

            if (rowOverride != null)
                row = rowOverride;

            for (int i = 0; i <= columns.Count - 1; i++)
            {
                string columnName = columns[i];

                if (columnName.EndsWith("_Dict_"))
                {
                    columnName = columnName.Replace("_Dict_", "");

                    object dictionaryObj = null;

                    if (sheetBuilder.GetCurrentTableName == columnName)
                    {
                        dictionaryObj = value;
                    } else if (columnName.Contains(":"))
                    {
                        dictionaryObj = GetFieldPathValue(value, columnName);
                    } else
                    {
                        dictionaryObj = (Dictionary<int, double>)GetFieldOrPropertyValue(value, columnName);
                    }

                    if (dictionaryObj == null || string.IsNullOrEmpty(dictionaryObj.ToString()))
                        continue;

                    object dictionaryKeys = dictionaryObj.GetType().GetProperty("Keys").GetValue(dictionaryObj);

                    int colCount = 0;
                    foreach (var key in (System.Collections.IEnumerable)dictionaryKeys)
                    {
                        ExcelCell keyCell = new ExcelCell();
                        keyCell.CellHeader = columnName + $":Key:{colCount}";
                        keyCell.CellValue = key;
                        ExcelColumnInfo info = null;
                        if (columnInfo != null)
                        {
                            info = columnInfo[i];
                            CreateReferenceCell(info, columnName, document, ref keyCell);
                        }
                        row.Add(keyCell);

                        var currentItem = dictionaryObj.GetType().GetProperty("Item").GetValue(dictionaryObj, new object[] { key });

                        if (FormatterUtils.IsSimpleType(currentItem.GetType()))
                        {
                            ExcelCell valueCell = new ExcelCell();
                            valueCell.CellHeader = columnName + $":Value:{colCount}";
                            valueCell.CellValue = currentItem;
                            row.Add(valueCell);
                        } else
                        {
                            string path = columnName + $":Value:{colCount}";
                            ExcelColumnInfoCollection listInnerObjectColumnInfo = _columnResolver.GetExcelColumnInfo(currentItem.GetType(), currentItem, path, true);
                            PopulateRows(listInnerObjectColumnInfo.Keys.ToList(), currentItem, sheetBuilder, listInnerObjectColumnInfo, document, row);
                        }



                        colCount++;
                    }
                } else if (columnName.EndsWith("_List_"))
                {
                    columnName = columnName.Replace("_List_", "");

                    List<object> listValue = new List<object>();

                    if (columnName.Contains(":"))
                    {
                        var valueObject = GetFieldPathValue(value, columnName);
                        if (valueObject != null && string.IsNullOrEmpty(valueObject.ToString()) == false)
                        {
                            listValue = (List<object>)valueObject;
                        }
                    } else
                    {
                        listValue = (List<object>)GetFieldOrPropertyValue(value, columnName);
                    }

                    int colCount = 0;
                    foreach (var kv in listValue)
                    {
                        string listColumnPrefix = columnName + $":{colCount}";

                        if (FormatterUtils.IsSimpleType(kv.GetType()))
                        {
                            ExcelCell listValueCell = new ExcelCell();
                            listValueCell.CellHeader = listColumnPrefix;
                            listValueCell.CellValue = kv;
                            row.Add(listValueCell);
                        } else
                        {
                            ExcelColumnInfoCollection listInnerObjectColumnInfo = _columnResolver.GetExcelColumnInfo(kv.GetType(), kv, listColumnPrefix, true);
                            PopulateRows(listInnerObjectColumnInfo.Keys.ToList(), kv, sheetBuilder, listInnerObjectColumnInfo, document, row);
                        }

                        colCount++;
                    }

                } else if (columnName.EndsWith("_CustomField_") || columnName.EndsWith("_CustomField_Single_"))
                {
                    bool isSingleValue = columnName.Contains("Single_");

                    columnName = columnName.Replace("_CustomField_", "").Replace("Single_", "");

                    List<object> customFields = null;
                    if (columnName.Contains(":"))
                    {
                        var valueObject = GetFieldPathValue(value, columnName);
                        if (valueObject != null && string.IsNullOrEmpty(valueObject.ToString()) == false)
                        {
                            customFields = (List<object>)valueObject;
                        }
                    } else
                    {
                        customFields = (List<object>)GetFieldOrPropertyValue(value, columnName);
                    }

                    if (customFields == null)
                        continue;

                    //need to get all custom columns

                    List<ExcelColumnInfo> allCustomColumns = null;
                    if (isSingleValue)
                    {
                        allCustomColumns = sheetBuilder.SheetColumns.Where(w => w.PropertyName == columnName).ToList();
                    } else
                    {
                        allCustomColumns = sheetBuilder.SheetColumns.Where(w => w.PropertyName.StartsWith(columnName)).ToList();
                    }


                    var objID = GetFieldOrPropertyValue(value, "ID");

                    foreach (var customColumn in allCustomColumns)
                    {
                        object objectCustomField = customFields.Where(w => customColumn.PropertyName.EndsWith($":{((dynamic)w).ID}")).Where(w => customColumn.PropertyName.Contains("Actual") ? ((dynamic)w).Actual == true : ((dynamic)w).Actual == false).FirstOrDefault();

                        ExcelCell customValueHeaderCell = new ExcelCell();

                        if (objectCustomField == null && !isSingleValue)
                        {
                            customValueHeaderCell.IsLocked = true;
                            customValueHeaderCell.CellHeader = customColumn.Header;
                            customValueHeaderCell.CellValue = "n/a";
                        } else
                        {
                            dynamic customFieldItem = (dynamic)objectCustomField;



                            string isActualText = string.Empty;
                            string columnNameCombined = string.Empty;

                            if (isSingleValue)
                            {
                                isActualText = columnName.Contains("Actual") ? ":Actual" : string.Empty;
                                columnNameCombined = $"{columnName}{isActualText}";
                                customFieldItem = (dynamic)customFields.First();
                            } else
                            {
                                isActualText = customFieldItem.Actual ? ":Actual" : string.Empty;
                                columnNameCombined = $"{columnName}{isActualText}:{customFieldItem.ID}";
                            }

                            customValueHeaderCell.CellHeader = customColumn.Header;

                            if (customFieldItem is CustomFieldModel)
                            {
                                if (!String.IsNullOrEmpty((customFieldItem as CustomFieldModel).Key))
                                {
                                    ExcelCell keyPreservationCell = new ExcelCell();
                                    keyPreservationCell.CellHeader = $"{columnNameCombined}:Key:{objID}";
                                    keyPreservationCell.CellValue = customFieldItem.Key?.ToString();
                                    CreatePreserveCell(keyPreservationCell, document);
                                }
                            }

                            ExcelCell valuePreservationCell = new ExcelCell();
                            valuePreservationCell.CellHeader = $"{columnNameCombined}:Value:{objID}";
                            if (customFieldItem != null)
                            {
                                valuePreservationCell.CellValue = customFieldItem.Value;
                                customValueHeaderCell.CellValue = customFieldItem.Value;

                                if (customFieldItem.Override != null)
                                {
                                    customValueHeaderCell.CellValue = customFieldItem.Override;
                                }
                            }


                            if (valuePreservationCell.CellValue != null && valuePreservationCell.CellValue.GetType() == typeof(DateTime))
                            {
                                valuePreservationCell.CellValue = valuePreservationCell.CellValue.ToString();
                            }
                            if (customValueHeaderCell.CellValue != null && customValueHeaderCell.CellValue.GetType() == typeof(DateTime))
                            {
                                customValueHeaderCell.CellValue = customValueHeaderCell.CellValue.ToString();
                            }

                            CreatePreserveCell(valuePreservationCell, document);

                            ExcelCell overridePreservationCell = new ExcelCell();
                            overridePreservationCell.CellHeader = $"{columnNameCombined}:Override:{objID}";
                            overridePreservationCell.CellValue = customFieldItem.Override?.ToString();
                            CreatePreserveCell(overridePreservationCell, document);

                            ExcelCell textPreservationCell = new ExcelCell();
                            textPreservationCell.CellHeader = $"{columnNameCombined}:Text:{objID}";
                            textPreservationCell.CellValue = customFieldItem.Text;
                            CreatePreserveCell(textPreservationCell, document);

                            ExcelCell hiddenTextPreservationCell = new ExcelCell();
                            hiddenTextPreservationCell.CellHeader = $"{columnNameCombined}:HiddenText:{objID}";
                            hiddenTextPreservationCell.CellValue = $"\"{customFieldItem.HiddenText}\"";
                            CreatePreserveCell(hiddenTextPreservationCell, document);

                            try
                            {
                                if (customFieldItem is CustomFieldModel
                                      && (customFieldItem as CustomFieldModel).ValueMix != null)
                                {
                                    foreach (var valueM in (customFieldItem as CustomFieldModel).ValueMix)
                                    {
                                        string valueMKey = valueM.Key;
                                        decimal valueMValue = valueM.Value;
                                        ExcelCell mixedPropertyPreservationCell = new ExcelCell();
                                        mixedPropertyPreservationCell.CellHeader = $"{columnNameCombined}:ValueMix:{valueM.Key}:{objID}";
                                        mixedPropertyPreservationCell.CellValue = $"\"{valueM.Value.ToString()}\"";
                                        CreatePreserveCell(mixedPropertyPreservationCell, document);
                                    }
                                }
                            } catch (Exception ex)
                            {
                                string s = ex.Message;
                            }

                            ExcelCell commonPropertyPreservationCell = new ExcelCell();
                            commonPropertyPreservationCell.CellHeader = $"{columnNameCombined}:Common:{objID}";
                            commonPropertyPreservationCell.CellValue = $"\"{customFieldItem.Common}\"";
                            CreatePreserveCell(commonPropertyPreservationCell, document);
                        }
                        row.Add(customValueHeaderCell);
                    }
                } else
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

            if (row.Count() > 0 && rowOverride == null)
                sheetBuilder.AppendRow(row.ToList());
        }

        private void PopulateReferenceSheet(XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder referenceSheet, DataTable ReferenceSheet)
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
                    } else if (c.DataType == typeof(DateTime))
                    {
                        if (r[c] is System.DBNull)
                            resolveRow.Add(c.Caption, string.Empty);
                        else
                            resolveRow.Add(c.Caption, Convert.ToDateTime(r[c]));
                    } else
                        resolveRow.Add(c.Caption, r[c].ToString());
                }

                PopulateRows(resolveRow.Keys.ToList(), resolveRow, referenceSheet);
            }

            //return sb;
        }

        private void PopulateInnerObjectSheets(ExcelSheetInfoCollection sheetsInfo, IXlsxDocumentBuilder document, Type itemType)
        {
            if (sheetsInfo == null)
                return;

            foreach (var sheet in sheetsInfo)
            {
                if (!(sheet.ExcelSheetAttribute is ExcelSheetAttribute))
                    continue;

                string sheetName = sheet.ExcelSheetAttribute != null ? (sheet.ExcelSheetAttribute as ExcelSheetAttribute).SheetName : itemType.Name;
                if (sheetName == null)
                    sheetName = sheet.SheetName;

                //sheetBuilder = new SqadXlsxSheetBuilder(document.AppendSheet(sheetName));

                this.Serialize(sheet.SheetType, sheet.SheetObject, document, sheetName);//, sheetBuilder);

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
            } else
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

                int tempValue = 0;
                if (int.TryParse(objName, out tempValue))
                    continue;

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
                    } else if ((matchingProperties = type.GetProperties().Where(w => w.Name == "Value").ToList()).Count() > 1)
                    {
                        //property overwriten, and must take first
                        member = matchingProperties.First();
                    } else
                    {
                        member = type.GetField(objName) ?? type.GetProperty(objName) as System.Reflection.MemberInfo;
                    }
                    if (member == null)
                    {
                        itemsToProcess = new List<object>() { rowObject };
                        continue;
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

                            int count = (int)result.GetType().GetProperty("Count").GetValue(result, null);

                            if (count == 0)
                                continue;

                            if (result.GetType().FullName.Contains("CustomFieldModel") || result.GetType().FullName.Contains("OverrideProperty"))
                            {
                                List<int> ids = resultsList.Select(s => (int)((dynamic)s).ID).ToList();
                                var filtereCustomFields = (result as IEnumerable<object>).Where(w => ids.Contains(((dynamic)w).ID) == false).ToList();
                                resultsList.AddRange(filtereCustomFields);
                                isResultCustomField = true;
                            } else
                            {
                                foreach (var resultItem in (System.Collections.IList)result)
                                {
                                    resultsList.Add(resultItem);
                                }

                                isResultList = true;
                            }
                        } else if (result.GetType().FullName.Contains("CustomFieldModel") || result.GetType().FullName.Contains("OverrideProperty"))
                        {
                            var exist = resultsList.Any(a => ((dynamic)a).ID == ((dynamic)result).ID);
                            if (!exist)
                                resultsList.Add(result);
                            isResultCustomField = true;
                        } else if (result.GetType().Name.StartsWith("Dictionary"))
                        {
                            try
                            {
                                if (resultsList.Any())
                                {
                                    //TODO: Change this to dynamic/generic implementation, because this dictionary can possibly be not only 'Dictionary<int, double>'
                                    //This fix was done according to https://gitlab.com/SQAD-MT/Web/WEBPro/-/issues/3410
                                    var firstListItem = resultsList[0] as Dictionary<int, double>;
                                    var сastedResult = result as Dictionary<int, double>;

                                    foreach (var item in сastedResult)
                                    {
                                        firstListItem[item.Key] = item.Value;
                                    }
                                } else
                                {
                                    //if dictionary with outher values it will failure
                                    resultsList.Add(result);
                                }
                                isResultDictionary = true;
                            } catch
                            {

                            }
                        } else
                            resultsList.Add(result);
                    }
                }
            }

            string returnString = string.Empty;


            if (isResultCustomField)
            {
                return resultsList;
            } else if (isResultList)
            {
                return resultsList;
            } else if (isResultDictionary)
            {
                return resultsList.FirstOrDefault();
            } else
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

                else if (cellValue.GetType() == typeof(DateTime) || (!Double.TryParse(cellValue.ToString(), out var test1) && DateTime.TryParse(cellValue.ToString(), out var test)))
                    return string.Format("{0:MM/dd/yyyy}", DateTime.Parse(cellValue.ToString()));

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
            } else if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveFromTable) == false)
            {
                _resolveTables.TryGetValue(info.ExcelColumnAttribute.ResolveFromTable, out columntResolveTable);
                if (columntResolveTable == default(DataTable))
                {
                    columntResolveTable = _staticValuesResolver.GetRecordsByTableName(info.ExcelColumnAttribute.ResolveFromTable);
                    if (columntResolveTable != null)
                    {
                        _resolveTables.Add(info.ExcelColumnAttribute.ResolveFromTable, columntResolveTable);
                    }
                }
            }

            if (!_isExportJsonToXls)
            {
                if (columntResolveTable != null)
                {
                    columntResolveTable.TableName = info.ExcelColumnAttribute.ResolveFromTable;
                    if (string.IsNullOrEmpty(info.ExcelColumnAttribute.OverrideResolveTableName) == false)
                        columntResolveTable.TableName = info.ExcelColumnAttribute.OverrideResolveTableName;

                    cell.DataValidationSheet = columntResolveTable.TableName;

                    var referenceSheet = document.GetReferenceSheet() as XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder;

                    if (referenceSheet == null)
                    {
                        referenceSheet = new XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder(cell.DataValidationSheet, true);
                        document.AppendSheet(referenceSheet);
                    } else
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
                } else if (string.IsNullOrEmpty(info.ExcelColumnAttribute.ResolveFromTable) == false)
                {
                    columntResolveTable = _staticValuesResolver.GetRecordsByTableName(info.ExcelColumnAttribute.ResolveFromTable);
                }
            }
        }

        public void CreatePreserveCell(ExcelCell cell, IXlsxDocumentBuilder document)
        {
            string _PreservationSheetName_ = "PreservationSheet";

            var preservationSheet = document.GetPreservationSheet() as XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder;

            if (preservationSheet == null)
            {
                preservationSheet = new XlsxExportImport.Base.Builders.SqadXlsxSheetBuilder(_PreservationSheetName_, false, true, true);
                document.AppendSheet(preservationSheet);
            } else
            {
                preservationSheet.AddAndActivateNewTable(_PreservationSheetName_);
            }

            preservationSheet.AppendColumnHeaderRowItem(cell.CellHeader);


            Dictionary<string, object> resolveRow = new Dictionary<string, object>();

            resolveRow.Add(cell.CellHeader, cell.CellValue);

            PopulateRows(resolveRow.Keys.ToList(), resolveRow, preservationSheet);
        }
    }
}