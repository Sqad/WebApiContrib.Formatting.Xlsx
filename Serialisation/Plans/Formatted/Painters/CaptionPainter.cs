using System;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class CaptionPainter
    {
        private const int ROW_MULTIPLIER = 3;

        private readonly ExcelWorksheet _worksheet;
        private readonly int _rowsOffset;
        private readonly Dictionary<DateTime, int> _columnsLookup;

        public CaptionPainter(ExcelWorksheet worksheet, int rowsOffset, Dictionary<DateTime, int> columnsLookup)
        {
            _worksheet = worksheet;
            _rowsOffset = rowsOffset;
            _columnsLookup = columnsLookup;
        }

        public int DrawCaption(Text caption)
        {
            var startColumnIndex = _columnsLookup[caption.StartDate.AddDays(-1).Date];
            var endColumnIndex = _columnsLookup[caption.EndDate.AddDays(-1).Date];

            var startRowIndex = caption.RowStart * ROW_MULTIPLIER + _rowsOffset - 3;
            var endRowIndex = caption.RowEnd * ROW_MULTIPLIER + _rowsOffset;

            var shape = _worksheet.Drawings.AddShape(caption.ID.ToString(), eShapeStyle.Rect);

            shape.From.Row = startRowIndex;
            shape.From.Column = startColumnIndex;
            shape.To.Row = endRowIndex;
            shape.To.Column = endColumnIndex;

            shape.Text = caption.TextValue;

            var appearance = AppearanceHelper.GetAppearance(caption.Appearance);
            FormatCaption(shape, appearance);

            return endRowIndex;
        }

        private static void FormatCaption(ExcelShape captionShape, CellsAppearance appearance)
        {
            captionShape.Fill.Transparancy = 30;

            captionShape.TextAlignment = appearance.TextAlignment;
            captionShape.TextAnchoring = appearance.TextVerticalAlignment;

            captionShape.Font.Color = appearance.TextColor;
            captionShape.Font.Size = appearance.FontSize;

            captionShape.Font.Bold = appearance.Bold;
            captionShape.Font.Italic = appearance.Italic;
            captionShape.Font.UnderLine = appearance.Underline ? eUnderLineType.Single : eUnderLineType.None;

            captionShape.Fill.Color = appearance.BackgroundColor;
            captionShape.Border.Fill.Color = appearance.CellBorderColor;
        }
    }
}