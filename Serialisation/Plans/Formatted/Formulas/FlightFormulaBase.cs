using System;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Formulas
{
    internal abstract class FlightFormulaBase
    {
        public abstract Func<Flight, object> GetValueAccessor(string formula);
    }
}
