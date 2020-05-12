using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers
{
    public class FlightHelper
    {
        public FlightHelper(Flight flight, int startCorrection = 0, int endCorrection = 0)
        {
            Flight = flight;
            StartCorrection = startCorrection;
            EndCorrection = endCorrection;
        }

        public Flight Flight { get; }
        public int StartCorrection { get; set; }
        public int EndCorrection { get; set; }

    }
}
