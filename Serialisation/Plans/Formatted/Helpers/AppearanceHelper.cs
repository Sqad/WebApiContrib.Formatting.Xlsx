using System.Drawing;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers
{
    internal static class AppearanceHelper
    {
        public static CellsAppearance GetAppearance(Flight flight, VehicleModel vehicle)
        {
            var appearance = GetMergedAppearance(flight.Appearance, vehicle?.Appearance);
            var cellsAppearance = new CellsAppearance();

            if (appearance.UseBackColor ?? false)
            {
                cellsAppearance.BackgroundColor = ColorTranslator.FromHtml(appearance.BackColor);
            }
            else
            {
                cellsAppearance.BackgroundColor = Colors.DefaultFlightBackgroundColor;
            }

            if (appearance.UseCellBorderColor ?? false)
            {
                cellsAppearance.CellBorderColor = ColorTranslator.FromHtml(appearance.CellBorderColor);
            }
            else
            {
                cellsAppearance.CellBorderColor = cellsAppearance.BackgroundColor;
            }

            return cellsAppearance;
        }

        private static Appearance GetMergedAppearance(Appearance childAppearance, Appearance baseAppearance)
        {
            var appearance = new Appearance();

            if (baseAppearance != null)
            {
                appearance.UseBackColor = baseAppearance.UseBackColor;
                appearance.BackColor = baseAppearance.BackColor;

                appearance.UseCellBorderColor = baseAppearance.UseCellBorderColor;
                appearance.CellBorderColor = baseAppearance.CellBorderColor;
            }

            appearance.UseBackColor = childAppearance.UseBackColor ?? appearance.UseBackColor;
            appearance.BackColor = childAppearance.BackColor ?? appearance.BackColor;
            appearance.UseCellBorderColor = childAppearance.UseCellBorderColor ?? appearance.UseCellBorderColor;
            appearance.CellBorderColor = childAppearance.CellBorderColor ?? appearance.CellBorderColor;

            return appearance;
        }
    }

    public class CellsAppearance
    {
        public Color BackgroundColor { get; set; }
        public Color CellBorderColor { get; set; }
    }
}