﻿using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class SQADActualsXlsSerialiser : IXlsxSerialiser
    {
        public bool CanSerialiseType(Type valueType, Type itemType)
        {
            throw new NotImplementedException();
        }

        public void Serialise(Type itemType, object value, IXlsxDocumentBuilder document, string sheetName)
        {
            throw new NotImplementedException();
        }
    }
}
