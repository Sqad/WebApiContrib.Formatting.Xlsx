using System;
using SQAD.XlsxExportImport.Base.Builders;
using SQAD.XlsxExportImport.Base.Serialization;

namespace SQAD.XlsxExportImport.Base.Interfaces
{
    /// <summary>
    /// Exposes access to serialization logic for complete customization of serialized output.
    /// </summary>
    public interface IXlsxSerializer
    {
        SerializerType SerializerType { get; }

        /// <summary>
        /// If true, no formatting beyond auto-fitting rows should be applied after serialization.
        /// </summary>
        // bool IgnoreFormatting { get; }

        /// <summary>
        /// Indicates whether the provided types can be serialized by this serializer implementation.
        /// </summary>
        /// <param name="valueType">Type of the value passed for serialization.</param>
        /// <param name="itemType">Type of items being serialized if value implements <c>IEnumerable</c>. (Will be the
        /// same as <c>valueType</c> otherwise.)</param>
        /// <returns></returns>
        bool CanSerializeType(Type valueType, Type itemType);

        /// <summary>
        /// Serialize the 
        /// </summary>
        /// <param name="itemType">Type of item being serialized.</param>
        /// <param name="value">Value passed for serialization, cast to an <c>IEnumerable</c> if necessary.</param>
        /// <param name="document">Document builder utility class.</param>
        void Serialize(Type itemType,
                    object value,
                    IXlsxDocumentBuilder document,
                    string sheetName,
                    string columnPrefix,
                    SqadXlsxSheetBuilder sheetbuilderOverride);
    }
}