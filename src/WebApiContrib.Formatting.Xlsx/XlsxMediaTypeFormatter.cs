using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Security.Permissions;
using System.Threading.Tasks;
using System.Data;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.Business.Models.Attributes;
using SQAD.MTNext.Services.Repositories.Export;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx
{

    /// <summary>
    /// Class used to send an Excel file to the response.
    /// </summary>
    public class XlsxMediaTypeFormatter : MediaTypeFormatter
    {

        #region Properties

        /// <summary>
        /// An action method that can be used to set the default cell style.
        /// </summary>
        protected Action<ExcelStyle> CellStyle { get; set; }

        /// <summary>
        /// An action method that can be used to set the default header row style.
        /// </summary>
        protected Action<ExcelStyle> HeaderStyle { get; set; }

        /// <summary>
        /// True if columns should be auto-fit to the cell contents after writing.
        /// </summary>
        protected bool AutoFit { get; set; }

        /// <summary>
        /// True if an auto-filter should be enabled for the data.
        /// </summary>
        protected bool AutoFilter { get; set; }

        /// <summary>
        /// True if the header row should be frozen.
        /// </summary>
        protected bool FreezeHeader { get; set; }

        /// <summary>
        /// Height for header row. (Default if null.)
        /// </summary>
        protected double? HeaderHeight { get; set; }

        /// <summary>
        /// Non-default serialisers to be used by this formatter instance.
        /// </summary>
        public List<IXlsxSerialiser> Serialisers { get; set; }

        public DefaultXlsxSerialiser DefaultSerializer { get; set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Create a new ExcelMediaTypeFormatter.
        /// </summary>
        /// <param name="autoFit">True if the formatter should autofit columns after writing the data. (Default
        /// true.)</param>
        /// <param name="autoFilter">True if an autofilter should be applied to the worksheet. (Default false.)</param>
        /// <param name="freezeHeader">True if the header row should be frozen. (Default false.)</param>
        /// <param name="headerHeight">Height of the header row.</param>
        /// <param name="cellHeight">Height of each row of data.</param>
        /// <param name="cellStyle">An action method that modifies the worksheet cell style.</param>
        /// <param name="headerStyle">An action method that modifies the cell style of the first (header) row in the
        /// worksheet.</param>
        public XlsxMediaTypeFormatter(bool autoFit = true,
                                      bool autoFilter = false,
                                      bool freezeHeader = false,
                                      double? headerHeight = null,
                                      Action<ExcelStyle> cellStyle = null,
                                      Action<ExcelStyle> headerStyle = null,
                                      IExportHelpersRepository staticValuesResolver = null)
        {
            SupportedMediaTypes.Clear();
            SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.ms-excel"));

            AutoFit = autoFit;
            AutoFilter = autoFilter;
            FreezeHeader = freezeHeader;
            HeaderHeight = headerHeight;
            CellStyle = cellStyle;
            HeaderStyle = headerStyle;

            // Initialise serialisers.
            Serialisers = new List<IXlsxSerialiser> { new SQADPlanXlsSerialiser(staticValuesResolver) };

            //DefaultSerializer = new SqadXlsxSerialiser(staticValuesResolver); //new DefaultXlsxSerialiser();
        }

        #endregion

        #region Methods

        public override void SetDefaultContentHeaders(Type type,
                                                      HttpContentHeaders headers,
                                                      MediaTypeHeaderValue mediaType)
        {

            string fileName = "data";

            // Look for ExcelDocumentAttribute on class.
            var itemType = FormatterUtils.GetEnumerableItemType(type);
            var excelDocumentAttribute = FormatterUtils.GetAttribute<ExcelDocumentAttribute>(itemType ?? type);

            if (excelDocumentAttribute != null && !string.IsNullOrEmpty(excelDocumentAttribute.FileName))
            {
                // If attribute exists with file name defined, use that.
                fileName = excelDocumentAttribute.FileName;
            }
            else
            {
                // Get the raw request URI.
                string rawUri = System.Web.HttpContext.Current?.Request?.RawUrl;
                if (string.IsNullOrEmpty(rawUri) != false)
                {
                    // Remove query string if present.
                    int queryStringIndex = rawUri.IndexOf('?');
                    if (queryStringIndex > -1)
                    {
                        rawUri = rawUri.Substring(0, queryStringIndex);
                    }
                }

                // Otherwise, use either the URL file name component or just "data".
                fileName = System.Web.VirtualPathUtility.GetFileName(rawUri) ?? "data";
            }

            // Add XLSX extension if not present.
            if (!fileName.EndsWith("xlsm", StringComparison.CurrentCultureIgnoreCase)) fileName += ".xlsm";

            // Set content disposition to use this file name.
            headers.ContentDisposition = new ContentDispositionHeaderValue("inline")
            { FileName = fileName };

            base.SetDefaultContentHeaders(type, headers, mediaType);
        }

        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override Task WriteToStreamAsync(Type type,
                                                object value,
                                                System.IO.Stream writeStream,
                                                System.Net.Http.HttpContent content,
                                                System.Net.TransportContext transportContext)
        {
            // Create a document builder.
            var document = new SqadXlsxDocumentBuilder(writeStream);

            if (value == null)
                return document.WriteToStream();

            var valueType = value.GetType();



            // Get the item type.
            var itemType = (FormatterUtils.IsSimpleType(valueType))
                ? null
                : FormatterUtils.GetEnumerableItemType(valueType);

            // If a single object was passed, treat it like a list with one value.
            if (itemType == null)
            {
                itemType = valueType;
                //value = new object[] { value };
            }

            // Used if no other matching serialiser can be found.
            IXlsxSerialiser serialiser = null;// new SqadXlsxSerialiser(_staticValuesResolver); //DefaultSerializer;

            // Determine if a more specific serialiser might apply.
            foreach (var s in Serialisers)
            {
                if (s.CanSerialiseType(valueType, itemType))
                {
                    serialiser = s;
                    break;
                }
            }

            serialiser.Serialise(itemType, value, document, null);

            if (!document.IsVBA)
            {
                content.Headers.ContentDisposition.FileName = content.Headers.ContentDisposition.FileName.Replace("xlsm", "xlsx");
            }

            return document.WriteToStream();
        }

        /// <summary>
        /// Applies custom formatting to a document. (Used if matched serialiser supports formatting.)
        /// </summary>
        /// <param name="document">The <c>XlsxDocumentBuilder</c> wrapping the document to format.</param>
        private void FormatDocument(XlsxDocumentBuilder document)
        {
            // Header cell styles
            if (HeaderStyle != null) HeaderStyle(document.Worksheet.Row(1).Style);
            if (FreezeHeader) document.Worksheet.View.FreezePanes(2, 1);

            var cells = document.Worksheet.Cells[document.Worksheet.Dimension.Address];

            // Add autofilter and fit to max column width (if requested).
            if (AutoFilter) cells.AutoFilter = AutoFilter;
            if (AutoFit) cells.AutoFitColumns();

            // Set header row where specified.
            if (HeaderHeight.HasValue)
            {
                document.Worksheet.Row(1).Height = HeaderHeight.Value;
                document.Worksheet.Row(1).CustomHeight = true;
            }
        }

        public override bool CanWriteType(Type type)
        {
            // Should be able to serialise any type.
            return true;
        }

        public override bool CanReadType(Type type)
        {
            // Not yet possible; see issue page to track progress:
            // https://github.com/jordangray/xlsx-for-web-api/issues/5
            return false;
        }

        #endregion

    }
}
