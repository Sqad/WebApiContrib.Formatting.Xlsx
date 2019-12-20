using System;
using System.Collections.Generic;
using System.Net.Http.Formatting;
using System.Text;

namespace WebApiContrib.Formatting.Xlsx
{
    public class SqadXlsxMediaTypeTemplateFormatter : MediaTypeFormatter
    {
        public override bool CanReadType(Type type)
        {
            throw new NotImplementedException();
        }

        public override bool CanWriteType(Type type)
        {
            throw new NotImplementedException();
        }
    }
}
