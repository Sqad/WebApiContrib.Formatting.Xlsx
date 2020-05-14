using System;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class CaptionPainter
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly Dictionary<DateTime, int> _columnsLookup;
        private readonly Dictionary<int, RowDefinition> _planRows;

        public CaptionPainter(ExcelWorksheet worksheet,
                              Dictionary<DateTime, int> columnsLookup,
                              Dictionary<int, RowDefinition> planRows)
        {
            _worksheet = worksheet;
            _columnsLookup = columnsLookup;
            _planRows = planRows;
        }

        public void DrawCaption(Text caption)
        {
            var startColumnIndex = _columnsLookup[caption.StartDate.Date] - 1;
            var endColumnIndex = _columnsLookup[caption.EndDate.AddDays(-1).Date];

            var startRow = _planRows[caption.RowStart];
            var endRow = _planRows[caption.RowEnd];

            var startRowIndex = startRow.StartExcelRowIndex - 1;
            var endRowIndex = endRow.EndExcelRowIndex;

            var shape = _worksheet.Drawings.AddShape(caption.ID.ToString(), eShapeStyle.Rect);

            shape.From.Row = startRowIndex;
            shape.From.Column = startColumnIndex;
            shape.To.Row = endRowIndex;
            shape.To.Column = endColumnIndex;

            shape.Text = caption.TextValue;

            var appearance = AppearanceHelper.GetAppearance(caption.Appearance);
            FormatCaption(shape, appearance);
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