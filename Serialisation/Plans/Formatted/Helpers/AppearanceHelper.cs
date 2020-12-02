using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using SQAD.MTNext.Business.Models.Core.Currency;
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
            if (childAppearance == null)
            {
                childAppearance = new Appearance();
            }

            var appearance = GetMergedAppearance(childAppearance, parentAppearance);

            var cellsAppearance = new CellsAppearance();

            if (appearance.UseBackColor ?? false)
            {
                cellsAppearance.UseBackColor = true;
                cellsAppearance.BackgroundColor = ColorTranslator.FromHtml(appearance.BackColor);
            }
            else
            {
                cellsAppearance.UseBackColor = false;
                cellsAppearance.BackgroundColor = Colors.DefaultFlightBackgroundColor;
            }

            if (appearance.UseCellBorderColor ?? false)
            {
                cellsAppearance.UseCellBorderColor = true;
                cellsAppearance.CellBorderColor = ColorTranslator.FromHtml(appearance.CellBorderColor);
            }
            else
            {
                cellsAppearance.UseCellBorderColor = false;
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

            if (appearance.UseCurrencySymbol ?? false)
            {
                cellsAppearance.UseCurrencySymbol = true;
                cellsAppearance.CurrencySymbol = appearance.CurrencySymbol ?? 0;
                cellsAppearance.CurrencySymbolAlign = appearance.CurrencySymbolAlign;
            }
            else
            {
                cellsAppearance.UseCurrencySymbol = false;
            }

            cellsAppearance.UsePercent = appearance.UsePercent ?? false;
            cellsAppearance.DigitGroupingChar = appearance.DigitGroupingChar ?? ",";
            if (cellsAppearance.DigitGroupingChar == ".")
            {
                cellsAppearance.DigitGroupingChar = ",";
            }

            cellsAppearance.FloatingPointAccuracy = appearance.FloatingPointAccuracy ?? 0;
            cellsAppearance.UseImageFillSizing = appearance.UseImageFillSizing ?? false;

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

            target.UseCurrencySymbol = source.UseCurrencySymbol ?? target.UseCurrencySymbol;
            target.CurrencySymbol = source.CurrencySymbol ?? target.CurrencySymbol;
            target.CurrencySymbolAlign = source.CurrencySymbolAlign ?? target.CurrencySymbolAlign;

            target.UsePercent = source.UsePercent ?? target.UsePercent;
            target.DigitGroupingChar = source.DigitGroupingChar ?? target.DigitGroupingChar;
            target.FloatingPointAccuracy = source.FloatingPointAccuracy ?? target.FloatingPointAccuracy;
            target.UseImageFillSizing = source.UseImageFillSizing ?? target.UseImageFillSizing;
        }
    }

    public class CellsAppearance
    {
        private static readonly Dictionary<string, string> CurrencyCodes;

        static CellsAppearance()
        {
            CurrencyCodes = CultureInfo.GetCultures(CultureTypes.AllCultures)
                                       .Where(x => !x.IsNeutralCulture)
                                       .Select(x =>
                                               {
                                                   try
                                                   {
                                                       return new RegionInfo(x.Name);
                                                   }
                                                   catch
                                                   {
                                                       return null;
                                                   }
                                               })
                                       .Where(x => x != null)
                                       .GroupBy(x => x.ISOCurrencySymbol)
                                       .ToDictionary(x => x.Key,
                                                     x => x.First().CurrencySymbol,
                                                     StringComparer.InvariantCultureIgnoreCase);
        }

        public bool UseBackColor { get; set; }
        public Color BackgroundColor { get; set; }
        public bool UseCellBorderColor { get; set; }
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

        public bool UseCurrencySymbol { get; set; }
        public int CurrencySymbol { get; set; }
        public string CurrencySymbolAlign { get; set; }
        public bool UsePercent { get; set; }
        public string DigitGroupingChar { get; set; }
        public int FloatingPointAccuracy { get; set; }

        public bool UseImageFillSizing { get; set; }

        public void FillValue(object value,
                              ExcelRange range,
                              Dictionary<int, CurrencyModel> currencies,
                              bool keepEmptyValues = true)
        {
            range.Value = value;

            var numericValue = value as double? ?? value as long?;
            if (numericValue.HasValue)
            {
                range.Style.Numberformat.Format = GetFormatString(currencies, numericValue.Value, keepEmptyValues);
            }

            if (value is DateTime)
            {
                range.Style.Numberformat.Format = "mm-dd-yy";
            }
        }

        public string GetValue(object value,
                                Dictionary<int, CurrencyModel> currencies)
        {
            var numericValue = value as double? ?? value as long?;
            if (numericValue.HasValue)
            {
                var numberFormat = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
                numberFormat.NumberGroupSeparator = DigitGroupingChar;
                numberFormat.CurrencyGroupSeparator = DigitGroupingChar;
                numberFormat.PercentGroupSeparator = DigitGroupingChar;

                numberFormat.NumberDecimalDigits = FloatingPointAccuracy;
                numberFormat.CurrencyDecimalDigits = FloatingPointAccuracy;
                numberFormat.PercentDecimalDigits = FloatingPointAccuracy;

                if (UsePercent)
                {
                    return numericValue.Value.ToString("P");
                }

                if (UseCurrencySymbol)
                {
                    var currency = currencies.GetValueOrDefault(CurrencySymbol);
                    if (currency != null)
                    {
                        var symbol = currency.CurrencySymbol;
                        if (string.IsNullOrWhiteSpace(symbol) && !CurrencyCodes.TryGetValue(currency.Code, out symbol))
                        {
                            symbol = "";
                        }

                        var formatted = numericValue.Value.ToString("N");

                        return CurrencySymbolAlign == "right"
                                   ? $"{formatted}{symbol}"
                                   : $"{symbol}{formatted}";
                    }
                }

                return numericValue.Value.ToString("N");
            }

            if (value is DateTime dateTime)
            {
                return dateTime.ToString("MM/dd/yyyy");
            }

            return value.ToString();
        }

        private string GetFormatString(IReadOnlyDictionary<int, CurrencyModel> currencies,
                                       double value,
                                       bool keepEmptyValues)
        {
            var isValueEmpty = (long) value == 0;

            if (UsePercent)
            {
                return isValueEmpty && keepEmptyValues ? "0%" : $"#{DigitGroupingChar}###{GetFloating()}%";
            }

            if (UseCurrencySymbol && currencies != null)
            {
                if (isValueEmpty && !keepEmptyValues)
                {
                    return "#";
                }

                var currency = currencies.GetValueOrDefault(CurrencySymbol);
                if (currency != null)
                {
                    var symbol = currency.CurrencySymbol;
                    if (string.IsNullOrWhiteSpace(symbol) && !CurrencyCodes.TryGetValue(currency.Code, out symbol))
                    {
                        symbol = "";
                    }

                    var format = isValueEmpty 
                                     ? "0" 
                                     : $"#{DigitGroupingChar}###{GetFloating()}";

                    return CurrencySymbolAlign == "right"
                               ? $"{format}{symbol}"
                               : $"{symbol}{format}";
                }
            }

            return isValueEmpty && keepEmptyValues ? "0" : $"#{DigitGroupingChar}###{GetFloating()}";
        }

        private string GetFloating()
        {
            if (FloatingPointAccuracy <= 0)
            {
                return "";
            }

            var numbers = string.Join("", Enumerable.Repeat("0", FloatingPointAccuracy));
            return $"0.{numbers}";
        }
    }
}