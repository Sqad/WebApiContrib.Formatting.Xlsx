using System;
using System.Data;
using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewDataSheetBuilder : SqadXlsxSheetBuilderBase
    {
        public SqadXlsxUnformattedViewDataSheetBuilder(string sheetName, bool isReferenceSheet = false, bool shouldAutoFit = true)
            : base(sheetName, isReferenceSheet, shouldAutoFit)
        {
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            throw new NotImplementedException();
        }
    }
}
