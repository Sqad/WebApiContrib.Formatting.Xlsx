using System;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Formulas
{
    internal class CurrencyNameFormula : FlightFormulaBase
    {
        public override Func<Flight, object> GetValueAccessor(string formula)
        {
            throw new NotImplementedException();
        }
    }
}
