using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx
{
    public interface IXlsxDocumentBuilder
    {
        Task WriteToStream();
        bool IsExcelSupportedType(object expression);
    }
}
