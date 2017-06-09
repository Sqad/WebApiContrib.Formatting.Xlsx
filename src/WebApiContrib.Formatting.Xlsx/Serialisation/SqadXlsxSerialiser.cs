using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class SqadXlsxSerialiser : IXlsxSerialiser
    {
        private IColumnResolver _columnResolver { get; set; }
        private ISheetResolver _sheetResolver { get; set; }

        public bool IgnoreFormatting => throw new NotImplementedException();

        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            return true;
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document)
        {
            var data = value as IEnumerable<object>;
        }
    }
}
