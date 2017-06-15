﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web.ModelBinding;
using WebApiContrib.Formatting.Xlsx.Attributes;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
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
                var propName = field;
                var prop = properties.FirstOrDefault(p => p.Name == propName);

                if (prop == null) continue;


                ExcelColumnAttribute attribute = FormatterUtils.GetAttribute<ExcelColumnAttribute>(prop);
                if (attribute != null)
                {
                    if (!FormatterUtils.IsSimpleType(prop.PropertyType))
                    {
                        //getting a complex class columns populates as ComplexName:InnerProperty
                        ExcelColumnInfoCollection columnCollection = GetExcelColumnInfo(prop.PropertyType, null, prop.Name, true);
                        foreach (var subcolumn in columnCollection)
                            fieldInfo.Add(subcolumn);
                    }
                    else
                    {
                        string propertyName = isComplexColumn ? $"{namePrefix}:{field}" : field; 
                        fieldInfo.Add(new ExcelColumnInfo(propertyName, attribute, null));
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
                    where p.CanRead & p.GetGetMethod().IsPublic & p.GetGetMethod().GetParameters().Length == 0
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
