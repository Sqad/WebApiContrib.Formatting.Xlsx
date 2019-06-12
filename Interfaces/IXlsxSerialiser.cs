using System;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Plans;

namespace SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces
{
    /// <summary>
    /// Exposes access to serialisation logic for complete customisation of serialised output.
    /// </summary>
    public interface IXlsxSerialiser
    {
        SerializerType SerializerType { get; }

        /// <summary>
        /// If true, no formatting beyond auto-fitting rows should be applied after serialisation.
        /// </summary>
       // bool IgnoreFormatting { get; }

        /// <summary>
        /// Indicates whether the provided types can be serialised by this serialiser implementation.
        /// </summary>
        /// <param name="valueType">Type of the value passed for serialisation.</param>
        /// <param name="itemType">Type of items being serialised if value implements <c>IEnumerable</c>. (Will be the
        /// same as <c>valueType</c> otherwise.)</param>
        /// <returns></returns>
        bool CanSerialiseType(Type valueType, Type itemType);

        /// <summary>
        /// Serialise the 
        /// </summary>
        /// <param name="itemType">Type of item being serialised.</param>
        /// <param name="value">Value passed for serialisation, cast to an <c>IEnumerable</c> if necessary.</param>
        /// <param name="document">Document builder utility class.</param>
        void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName, string columnPrefix, SqadXlsxPlanSheetBuilder sheetbuilderOverride);


    }
}
