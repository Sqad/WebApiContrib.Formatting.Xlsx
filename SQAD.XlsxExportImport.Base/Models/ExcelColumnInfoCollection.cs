﻿namespace SQAD.XlsxExportImport.Base.Models
{
    /// <summary>
    /// A collection of column information for an Excel document, keyed by field/property name.
    /// </summary>
    public class ExcelColumnInfoCollection : KeyedCollectionBase<ExcelColumnInfo>
    {
        protected override string GetKeyForItem(ExcelColumnInfo item)
        {
            return item.PropertyName;
        }
    }
}
