using System;

namespace SQAD.XlsxExportImport.Base.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        // Nullable parameters not allowed on attributes. :(
        public int? _order;

        /// <summary>
        /// Control the output of this property when serialized to Excel.
        /// </summary>
        public ExcelColumnAttribute() { }

        /// <summary>
        /// Control the output of this property when serialized to Excel.
        /// </summary>
        public ExcelColumnAttribute(string header) : this()
        {
            Header = header;
        }

        /// <summary>
        /// Column header to use for this property.
        /// </summary>
        public string Header { get; set; }

        /// <summary>
        /// Value to use if this field is a boolean value and equals <c>true</c>.
        /// </summary>
        public string TrueValue { get; set; }

        /// <summary>
        /// Value to use if this field is a boolean value and equals <c>false</c>.
        /// </summary>
        public string FalseValue { get; set; }

        /// <summary>
        /// Whether to use the display format string set for this field.
        /// </summary>
        public bool UseDisplayFormatString { get; set; }

        /// <summary>
        /// Ignore this property when serializing to Excel.
        /// </summary>
        public bool Ignore { get; set; }

        /// <summary>
        /// Override the serialized order of this property in the generated Excel document.
        /// </summary>public int Order
        public int Order
        {
            get { return _order ?? default; }
            set { _order = value; }
        }

        /// <summary>
        /// Apply the specified Excel number format string to this property in the generated Excel output.
        /// </summary>
        public string NumberFormat { get; set; }

        public string ResolveFromTable { get; set; }

        public string OverrideResolveTableName { get; set; }
        /// <summary>
        /// Default Value is ID
        /// </summary>
        public string ResolveValue { get; set; } = "ID";
        /// <summary>
        /// Default Value is Name
        /// </summary>
        public string ResolveName { get; set; } = "Name";

        /// <summary>
        /// When we have subproperty that we want to skip name
        /// like ParentPropertyName:Value(this property). Value will not be included
        /// in column name path
        /// </summary>
        public bool DoNotUsePropertyName { get; set; } = false;

        public bool IsHidden { get; set; }

        /// <summary>
        /// Used to items containing override value. 
        /// Original value will be preserved and if changed will be moved to override property
        /// </summary>
        public bool PreserverHidden { get; set; }
    }
}
