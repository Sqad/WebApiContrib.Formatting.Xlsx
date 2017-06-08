using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class SqadXlsxSerialiser : DefaultXlsxSerialiser
    {
        public override void Serialise(Type itemType, object value, XlsxDocumentBuilder document)
        {
            base.Serialise(itemType, value, document);
        }
    }
}
