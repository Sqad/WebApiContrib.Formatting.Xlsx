using SQAD.MTNext.Business.Models.Attributes;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web.ModelBinding;

namespace SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation
{
    /// <summary>
    /// Resolves all public, parameterless properties of an object, respecting any <c>ExcelColumnAttribute</c>
    /// values.
    /// </summary>
    public class DefaultColumnResolver : IColumnResolver
    {
        /// <summary>
        /// Get the <c>ExcelColumnInfo</c> for all members of a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        public virtual ExcelColumnInfoCollection GetExcelColumnInfo(Type itemType, object data, string namePrefix = "", bool isComplexColumn = false)
        {
            var fields = GetSerialisableMemberNames(itemType, data);
            var properties = GetSerialisablePropertyInfo(itemType, data);

            var fieldInfo = new ExcelColumnInfoCollection();

            // Instantiate field names and fieldInfo lists with serialisable members.
            foreach (var field in fields)
            {
                var prop = properties.FirstOrDefault(p => p.Name == field);

                if (prop == null) continue;

                Type propertyType = prop.PropertyType;
                if (propertyType.IsGenericType &&
                    propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    propertyType = propertyType.GetGenericArguments()[0];
                }

                ExcelColumnAttribute attribute = FormatterUtils.GetAttribute<ExcelColumnAttribute>(prop);
                if (attribute != null)
                {
                    string prefix = string.IsNullOrEmpty(namePrefix) == false ? $"{namePrefix}:{prop.Name}" : prop.Name;

                    if (propertyType.Name.StartsWith("List"))
                    {
                        Type typeOfList = FormatterUtils.GetEnumerableItemType(propertyType);

                        if (FormatterUtils.IsSimpleType(typeOfList))
                        {
                            fieldInfo.Add(new ExcelColumnInfo(prefix, typeOfList, attribute, null));
                        }
                        else if (typeOfList.FullName.EndsWith("CustomFieldModel") || typeOfList.Name.StartsWith("OverrideProperty"))
                        {
                            prefix += "_CustomField_";
                            fieldInfo.Add(new ExcelColumnInfo(prefix, null, attribute, null));
                        }
                        else
                        {
                            prefix += "_List_";
                            fieldInfo.Add(new ExcelColumnInfo(prefix, null, attribute, null));

                            //ExcelColumnInfoCollection columnCollection = GetExcelColumnInfo(typeOfList, null, prefix, true);
                            //foreach (var subcolumn in columnCollection)
                            //    fieldInfo.Add(subcolumn);
                        }
                    }
                    else if (propertyType.Name.StartsWith("Dictionary"))
                    {
                        prefix += "_Dict_";
                        fieldInfo.Add(new ExcelColumnInfo(prefix, null, attribute, null));
                    }
                    else if (!FormatterUtils.IsSimpleType(propertyType))
                    {
                        ExcelColumnInfoCollection columnCollection = GetExcelColumnInfo(propertyType, null, prefix, true);
                        foreach (var subcolumn in columnCollection)
                            fieldInfo.Add(subcolumn);
                    }
                    else
                    {
                        string propertyName = isComplexColumn ? $"{namePrefix}:{field}" : field;
                        string displayName = propertyName;

                        if (attribute.DoNotUsePropertyName)
                            attribute.Header = namePrefix;

                        bool columnAlreadyadded = fieldInfo.Any(a => a.PropertyName == propertyName);
                        if (!columnAlreadyadded)
                        {
                            if (FormatterUtils.IsExcelSupportedType(propertyType))
                                fieldInfo.Add(new ExcelColumnInfo(propertyName, propertyType, attribute, null));
                            else
                                fieldInfo.Add(new ExcelColumnInfo(propertyName, typeof(string), attribute, null));
                        }
                    }
                }
            }

            PopulateFieldInfoFromMetadata(fieldInfo, itemType, data);

            return fieldInfo;
        }

        /// <summary>
        /// Get a list of all non-ignored public instance property names for a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        public virtual IEnumerable<string> GetSerialisableMemberNames(Type itemType, object data)
        {
            return FormatterUtils.GetMemberNames(itemType);
        }

        /// <summary>
        /// Get <c>PropertyInfo</c> for all public instance properties with parameterless get methods in a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        public virtual IEnumerable<PropertyInfo> GetSerialisablePropertyInfo(Type itemType, object data)
        {
            return (from p in itemType.GetProperties()
                    where p.CanRead
                    where p.GetGetMethod() == null || p.GetGetMethod().IsPublic & p.GetGetMethod().GetParameters().Length == 0
                    select p).ToList();
        }

        /// <summary>
        /// Populate missing or incomplete properties from model metadata.
        /// </summary>
        /// <param name="fieldInfo">The <c>ExcelColumnInfoCollection</c> to populate.</param>
        /// <param name="itemType">The type of item whose metadata this is being populated from.</param>
        /// <param name="data">The collection of values being serialised. (Not used, provided for use by derived
        /// types.)</param>
        protected virtual void PopulateFieldInfoFromMetadata(ExcelColumnInfoCollection fieldInfo,
                                                             Type itemType,
                                                             object data)
        {
            // Populate missing attribute information from metadata.
            var metadata = ModelMetadataProviders.Current.GetMetadataForType(null, itemType);

            if (metadata != null && metadata.Properties != null)
            {
                foreach (var modelProp in metadata.Properties)
                {
                    var propertyName = modelProp.PropertyName;

                    if (!fieldInfo.Contains(propertyName)) continue;

                    var field = fieldInfo[propertyName];

                    field.PropertyType = modelProp.ModelType;
                    if (field.PropertyType.IsGenericType &&
                        field.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        field.PropertyType = field.PropertyType.GetGenericArguments()[0];
                    }


                    var attribute = field.ExcelColumnAttribute;

                    if (!field.IsExcelHeaderDefined)
                        field.Header = modelProp.DisplayName ?? propertyName;

                    if (attribute != null && attribute.UseDisplayFormatString)
                        field.FormatString = modelProp.DisplayFormatString;
                   
                }
            }
        }
    }
}
