using System.Drawing;
using OfficeOpenXml.Drawing;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers
{
    internal static class AppearanceHelper
    {
        public static CellsAppearance GetAppearance(Flight flight, VehicleModel vehicle)
        {
            return GetAppearance(flight.Appearance, vehicle?.Appearance);
        }

        public static CellsAppearance GetAppearance(Appearance childAppearance, Appearance parentAppearance = null)
        {
            var appearance = GetMergedAppearance(childAppearance, parentAppearance);

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

            cellsAppearance.TextColor = string.IsNullOrWhiteSpace(appearance.TextColor)
                                            ? Color.Black
                                            : ColorTranslator.FromHtml(appearance.TextColor);

            const int baseFontSizeDiff = 16 - 11;
            if (string.IsNullOrWhiteSpace(appearance.FontSize))
            {
                cellsAppearance.FontSize = 16 - baseFontSizeDiff;
            }
            else
            {
                var sourceValue = appearance.FontSize.Replace("em", "");
                if (!double.TryParse(sourceValue, out var value))
                {
                    value = 1;
                }

                cellsAppearance.FontSize = (int) (value * (16 - baseFontSizeDiff));
            }

            cellsAppearance.Bold = appearance.Bold ?? false;
            cellsAppearance.Italic = appearance.Italic ?? false;
            cellsAppearance.Underline = appearance.Underline ?? false;

            const string flexStartValue = "flex-start";
            const string flexEndValue = "flex-end";

            switch (appearance.TextAlign)
            {
                case flexStartValue:
                    cellsAppearance.TextAlignment = eTextAlignment.Left;
                    break;
                case flexEndValue:
                    cellsAppearance.TextAlignment = eTextAlignment.Right;
                    break;
                default:
                    cellsAppearance.TextAlignment = eTextAlignment.Center;
                    break;
            }

            switch (appearance.TextVerticalAlign)
            {
                case flexStartValue:
                    cellsAppearance.TextVerticalAlignment = eTextAnchoringType.Top;
                    break;
                case flexEndValue:
                    cellsAppearance.TextVerticalAlignment = eTextAnchoringType.Bottom;
                    break;
                default:
                    cellsAppearance.TextVerticalAlignment = eTextAnchoringType.Center;
                    break;
            }

            if (appearance.UseFill ?? false)
            {
                cellsAppearance.UseFillColor = true;
                cellsAppearance.FillColor = ColorTranslator.FromHtml(appearance.Fill);
            }
            else
            {
                cellsAppearance.UseFillColor = false;
                cellsAppearance.FillColor = Colors.DefaultFlightBackgroundColor;
            }

            if (appearance.UseStroke ?? false)
            {
                cellsAppearance.StrokeColor = ColorTranslator.FromHtml(appearance.Stroke);
            }
            else
            {
                cellsAppearance.StrokeColor = Colors.DefaultFlightBackgroundColor;
            }

            cellsAppearance.StrokeWidth = appearance.StrokeWidth ?? 3;
            cellsAppearance.RotationAngle = appearance.RotationAngle ?? 0;

            if (appearance.UseOutlineColor ?? false)
            {
                cellsAppearance.UseOutlineColor = true;
                cellsAppearance.OutlineColor = ColorTranslator.FromHtml(appearance.OutlineColor);
            }
            else
            {
                cellsAppearance.UseOutlineColor = false;
                cellsAppearance.OutlineColor = Colors.DefaultFlightBackgroundColor;
            }

            cellsAppearance.Transparency = appearance.Transparency ?? 1;

            return cellsAppearance;
        }

        private static Appearance GetMergedAppearance(Appearance childAppearance, Appearance baseAppearance)
        {
            var appearance = new Appearance();

            if (baseAppearance != null)
            {
                FillAppearance(appearance, baseAppearance);
            }

            FillAppearance(appearance, childAppearance);

            return appearance;
        }

        private static void FillAppearance(Appearance target, Appearance source)
        {
            target.UseBackColor = source.UseBackColor ?? target.UseBackColor;
            target.BackColor = source.BackColor ?? target.BackColor;

            target.UseCellBorderColor = source.UseCellBorderColor ?? target.UseCellBorderColor;
            target.CellBorderColor = source.CellBorderColor ?? target.CellBorderColor;

            target.TextColor = source.TextColor ?? target.TextColor;
            target.FontSize = source.FontSize ?? target.FontSize;
            target.Bold = source.Bold ?? target.Bold;
            target.Italic = source.Italic ?? target.Italic;
            target.Underline = source.Underline ?? target.Underline;

            target.TextAlign = source.TextAlign ?? target.TextAlign;
            target.TextVerticalAlign = source.TextVerticalAlign ?? target.TextVerticalAlign;
            target.TextDirection = source.TextDirection ?? target.TextDirection;

            target.UseFill = source.UseFill ?? target.UseFill;
            target.Fill = source.Fill ?? target.Fill;

            target.UseStroke = source.UseStroke ?? target.UseStroke;
            target.Stroke = source.Stroke ?? target.Stroke;
            target.StrokeWidth = source.StrokeWidth ?? target.StrokeWidth;

            target.RotationAngle = source.RotationAngle ?? target.RotationAngle;

            target.UseOutlineColor = source.UseOutlineColor ?? target.UseOutlineColor;
            target.OutlineColor = source.OutlineColor ?? target.OutlineColor;

            target.Transparency = source.Transparency ?? target.Transparency;
        }
    }

    public class CellsAppearance
    {
        public Color BackgroundColor { get; set; }
        public Color CellBorderColor { get; set; }

        public Color TextColor { get; set; }
        public int FontSize { get; set; }

        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }

        public eTextAlignment TextAlignment { get; set; }
        public eTextAnchoringType TextVerticalAlignment { get; set; }

        public bool UseFillColor { get; set; }
        public Color FillColor { get; set; }

        public Color StrokeColor { get; set; }
        public int StrokeWidth { get; set; }

        public int RotationAngle { get; set; }

        public bool UseOutlineColor { get; set; }
        public Color OutlineColor { get; set; }

        public double Transparency { get; set; }
    }
}