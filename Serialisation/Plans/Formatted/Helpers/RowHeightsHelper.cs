using System;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using System.Collections.Generic;
using System.Linq;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers
{
    internal static class RowHeightsHelper
    {
        private const int ROWS_OFFSET = 3;

        public static void FillRowHeights(Dictionary<int, RowDefinition> planRows, ChartData chartData)
        {
            var flights = chartData.Objects.Flights ?? new List<Flight>();
            var vehicles = (chartData.Vehicles ?? new List<VehicleModel>()).ToDictionary(x => x.ID);

            var maxRowIndex = 0;
            foreach (var flight in flights)
            {
                if (flight.VehicleID.HasValue)
                {
                    var vehicle = vehicles.GetValueOrDefault(flight.VehicleID.Value);

                    if (vehicle != null)
                    {
                        if (flight.FlightCaption == null)
                        {
                            flight.FlightCaption = new FlightCaption();
                        }

                        if (flight.FlightCaption.Above == null)
                        {
                            flight.FlightCaption.Above = new List<FlightCaptionPosition>();
                        }

                        if (flight.FlightCaption.Below == null)
                        {
                            flight.FlightCaption.Below = new List<FlightCaptionPosition>();
                        }

                        if (flight.FlightCaption.Inside == null)
                        {
                            flight.FlightCaption.Inside = new List<FlightCaptionPosition>();
                        }

                        if (flight.OverwrittenCaptions == null)
                        {
                            flight.OverwrittenCaptions = new List<Guid>();
                        }

                        var above = (vehicle.FlightCaption?.Above ?? new List<FlightCaptionPosition>())
                            .ToDictionary(x => x.ID);
                        var below = (vehicle.FlightCaption?.Below ?? new List<FlightCaptionPosition>())
                            .ToDictionary(x => x.ID);
                        var inside = (vehicle.FlightCaption?.Inside ?? new List<FlightCaptionPosition>())
                            .ToDictionary(x => x.ID);

                        foreach (var overwrite in flight.OverwrittenCaptions)
                        {
                            above.Remove(overwrite);
                            below.Remove(overwrite);
                            inside.Remove(overwrite);
                        }

                        flight.FlightCaption.Above.AddRange(above.Values);
                        flight.FlightCaption.Below.AddRange(below.Values);
                        flight.FlightCaption.Inside.AddRange(inside.Values);
                    }
                }

                var flightRowIndex = flight.RowIndex ?? 0;

                if (!planRows.TryGetValue(flightRowIndex, out var rowDefinition))
                {
                    rowDefinition = new RowDefinition
                                    {
                                        OriginalRowIndex = flightRowIndex
                                    };
                    planRows.Add(flightRowIndex, rowDefinition);
                }

                var aboveCaptionsCount = flight.FlightCaption?.Above?.Count ?? 0;
                var belowCaptionsCount = flight.FlightCaption?.Below?.Count ?? 0;

                if (aboveCaptionsCount > rowDefinition.AboveCount)
                {
                    rowDefinition.AboveCount = aboveCaptionsCount;
                }

                if (belowCaptionsCount > rowDefinition.BelowCount)
                {
                    rowDefinition.BelowCount = belowCaptionsCount;
                }

                if (flightRowIndex > maxRowIndex)
                {
                    maxRowIndex = flightRowIndex;
                }
            }

            var maxFormulaIndex = chartData.Objects.Formulas?.Max(x => x.RowIndex) ?? 0;
            var maxShapeIndex = chartData.Objects.Shapes?.Max(x => x.RowEnd) ?? 0;
            var maxPictureIndex = chartData.Objects.Pictures?.Max(x => x.RowEnd) ?? 0;
            var maxTextIndex = chartData.Objects.Texts?.Max(x => x.RowEnd) ?? 0;
            var maxLeftTableCellIndex = (chartData.Cells ?? new List<TextValue>())
                                        .Select(x => new CellAddress(x.Key))
                                        .Where(x => x.IsFlightsTableAddress)
                                        .Max(x => x.RowIndex);

            maxRowIndex = GetMax(maxRowIndex,
                                 maxFormulaIndex,
                                 maxShapeIndex,
                                 maxPictureIndex,
                                 maxTextIndex,
                                 maxLeftTableCellIndex);

            var offset = ROWS_OFFSET;
            for (var currentRowIndex = 1; currentRowIndex <= maxRowIndex; currentRowIndex++)
            {
                var rowDefinition = planRows.GetValueOrDefault(currentRowIndex);
                if (rowDefinition == null)
                {
                    var index = currentRowIndex + offset;
                    rowDefinition = new RowDefinition
                                    {
                                        OriginalRowIndex = currentRowIndex,
                                        PrimaryExcelRowIndex = index,
                                        StartExcelRowIndex = index,
                                        EndExcelRowIndex = index
                                    };
                    planRows.Add(currentRowIndex, rowDefinition);

                    continue;
                }

                rowDefinition.StartExcelRowIndex = rowDefinition.OriginalRowIndex + offset;
                rowDefinition.PrimaryExcelRowIndex = rowDefinition.StartExcelRowIndex + rowDefinition.AboveCount;
                rowDefinition.EndExcelRowIndex = rowDefinition.PrimaryExcelRowIndex + rowDefinition.BelowCount;

                offset += rowDefinition.AboveCount + rowDefinition.BelowCount;
            }
        }

        private static int GetMax(params int[] indexes)
        {
            return indexes.Max();
        }
    }
}
