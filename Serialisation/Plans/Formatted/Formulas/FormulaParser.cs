using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using Microsoft.EntityFrameworkCore.Internal;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Formulas
{
    internal class FormulaParser
    {
        private readonly PropertyFormula _propertyFormula;

        private readonly Dictionary<string, string> _fieldsMap = new Dictionary<string, string>
                                                                 {
                                                                     {"Row", "RowIndex"},
                                                                     {"Duration", "Days"}
                                                                 };

        public FormulaParser()
        {
            _propertyFormula = new PropertyFormula();
        }

        public string GetInsideCaption(Flight flight)
        {
            if (flight.FlightCaption.Inside == null
                || !flight.FlightCaption.Inside.Any())
            {
                return string.Empty;
            }

            var captionBuilder = new StringBuilder();

            foreach (var flightCaption in flight.FlightCaption.Inside)
            {
                captionBuilder.Append(GetValueFromFormula(flightCaption.Text, flight));
                captionBuilder.Append(" ");
            }

            return captionBuilder.ToString();
        }

        public string GetValueFromFormula(string formula, Flight flight)
        {
            if (!formula.StartsWith("=Flight."))
            {
                return formula;
            }

            var propertyName = formula.Substring(8);
            return GetFlightPropertyByName(propertyName, flight);
        }

        private string GetFlightPropertyByName(string propertyName, Flight flight)
        {
            var propertyAccessor = _propertyFormula.GetValueAccessor(propertyName);
            if (propertyAccessor == null)
            {
                return GetFlightMeasure(propertyName, flight);
            }

            var value = propertyAccessor(flight);
            if (value is DateTime datetimeValue)
            {
                return datetimeValue.ToString("MM/dd/yyyy");
            }

            return value.ToString();
        }

        private string GetFlightMeasure(string propertyName, Flight flight)
        {
            var property = flight.Measures
                                 .GetType()
                                 .GetProperty(propertyName);
            if (property == null)
            {
                return string.Empty;
            }

            var value = property.GetValue(flight.Measures, null);

            property = value.GetType().GetProperty("ComputedValue");
            if (property == null)
            {
                return string.Empty;
            }

            value = property.GetValue(value, null);
            if (value is DateTime datetimeValue)
            {
                return datetimeValue.ToString("MM/dd/yyyy");
            }

            return value.ToString();
        }
    }
}