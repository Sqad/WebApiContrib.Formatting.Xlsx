using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Formulas
{
    internal class PropertyFormula : FlightFormulaBase
    {
        private readonly Dictionary<string, string> _fieldsMap = new Dictionary<string, string>
                                                                 {
                                                                     {nameof(Flight.Days), "Duration"},
                                                                     {nameof(Flight.RowIndex), "Row"}
                                                                 };

        private readonly Dictionary<string, Func<Flight, object>> _propertyCache;

        public PropertyFormula()
        {
            _propertyCache = CreatePropertyCache();
        }

        public override Func<Flight, object> GetValueAccessor(string formula)
        {
            return _propertyCache.GetValueOrDefault(formula);
        }

        private Dictionary<string, Func<Flight, object>> CreatePropertyCache()
        {
            var result = new Dictionary<string, Func<Flight, object>>();

            var flightParameter = Expression.Parameter(typeof(Flight));

            var properties = typeof(Flight).GetProperties();
            foreach (var property in properties)
            {
                var accessor = Expression.Convert(Expression.Property(flightParameter, property.Name), typeof(object));
                var lambda = Expression.Lambda<Func<Flight, object>>(accessor, flightParameter);

                var propertyName = property.Name;
                if (_fieldsMap.TryGetValue(propertyName, out var mappedName))
                {
                    propertyName = mappedName;
                }

                result.Add(propertyName, lambda.Compile());
            }

            return result;
        }
    }
}