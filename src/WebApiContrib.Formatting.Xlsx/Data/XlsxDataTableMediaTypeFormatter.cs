using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Security.Permissions;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Data
{
    public class XlsxDataTableMediaTypeFormatter : MediaTypeFormatter
    {
        public XlsxDataTableMediaTypeFormatter()
        {
            SupportedMediaTypes.Clear();
            SupportedMediaTypes
                .Add(new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            SupportedMediaTypes.Add(new MediaTypeHeaderValue("application/vnd.ms-excel"));
        }

        public override void SetDefaultContentHeaders(Type type,
                                                      HttpContentHeaders headers,
                                                      MediaTypeHeaderValue mediaType)
        {
            headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                                         {
                                             FileName = "test.xlsx"
                                         };

            base.SetDefaultContentHeaders(type, headers, mediaType);
        }

        [SecurityPermission(SecurityAction.Demand, SerializationFormatter = true)]
        public override Task WriteToStreamAsync(Type type,
                                                object value,
                                                Stream writeStream,
                                                HttpContent content,
                                                TransportContext transportContext)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");

            if (value == null || !(value is DataTable dataTable))
            {
                return Task.Factory.StartNew(() =>
                                             {
                                                 package.SaveAs(writeStream);
                                             });
            }

            var builder = new SqadDataTableSheetBuilder(worksheet);
            builder.BuildSheet(dataTable);

            return Task.Factory.StartNew(() =>
                                         {
                                             package.SaveAs(writeStream);
                                         });
        }

        public override bool CanReadType(Type type)
        {
            return false;
        }

        public override bool CanWriteType(Type type)
        {
            return type == typeof(DataTable);
        }
    }
}