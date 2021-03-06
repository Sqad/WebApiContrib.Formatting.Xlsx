﻿using SQAD.XlsxExportImport.Base.Models;
using System;

namespace SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces
{
    /// <summary>
    /// Allows easy customisation of what columns are generated from a type and how they should be formatted.
    /// </summary>
    /// <remarks>Used by
    /// the <c>DefaultXlsxSerialiser</c> and derived types. See <c>DefaultColumnResolver</c> for a good starting
    /// point to write your own.</remarks>
    public interface IColumnResolver
    {
        /// <summary>
        /// Get the <c>ExcelColumnInfo</c> for all serialisable members of a class.
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="data">The collection of values being serialised.</param>
        //ExcelColumnInfoCollection GetExcelColumnInfo(Type itemType, IEnumerable<object> data);
        ExcelColumnInfoCollection GetExcelColumnInfo(Type itemType, object data, string namePrefix = "", bool isComplexColumn=false);
    }
}
