﻿using SQAD.MTNext.Business.Models.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using WebApiContrib.Formatting.Xlsx.Interfaces;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class DefaultSheetResolver : ISheetResolver
    {
        public ExcelSheetInfoCollection GetExcelSheetInfo(Type itemType, object data)
        {
            var sheets = FormatterUtils.GetMemberNames(itemType);
            var properties = GetSerialisablePropertyInfo(itemType);

            var sheetCollection = new  ExcelSheetInfoCollection();

            foreach (var sheet in sheets)
            {
                var prop = properties.FirstOrDefault(p => p.Name == sheet);

                ExcelSheetAttribute attr = FormatterUtils.GetAttribute<ExcelSheetAttribute>(prop);

                if (prop==null || attr==null ) continue;

                var sheetInfo = new ExcelSheetInfo()
                {
                    SheetType = prop.PropertyType,
                    SheetName = sheet,
                    ExcelSheetAttribute = attr,
                    PropertyName = prop.Name,
                    SheetObject = FormatterUtils.GetFieldOrPropertyValue(data, prop.Name)
                };

                if (prop.PropertyType.Name.StartsWith("List"))
                    sheetInfo.SheetType = FormatterUtils.GetEnumerableItemType(sheetInfo.SheetObject.GetType());

                sheetCollection.Add(sheetInfo);
            }

            return sheetCollection;
        }

        public virtual IEnumerable<PropertyInfo> GetSerialisablePropertyInfo(Type itemType)
        {
            return (from p in itemType.GetProperties()
                    where p.CanRead & p.GetGetMethod().IsPublic & p.GetGetMethod().GetParameters().Length == 0
                    select p).ToList();
        }
    }
}
