using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class ShapePainter
    {
        private const string LineType = "SHAPE_LINE";

        private static readonly Dictionary<string, eShapeStyle> ShapesMap;

        private readonly ExcelWorksheet _worksheet;
        private readonly Dictionary<DateTime, int> _columnsLookup;
        private readonly Dictionary<int, RowDefinition> _planRows;

        static ShapePainter()
        {
            ShapesMap = new Dictionary<string, eShapeStyle>(StringComparer.InvariantCultureIgnoreCase)
                        {
                            {"SHAPE_BOX", eShapeStyle.RoundRect},
                            {"SHAPE_CIRCLE", eShapeStyle.Ellipse},
                            {"SHAPE_ELLIPSE", eShapeStyle.Ellipse},
                            {"SHAPE_STAR", eShapeStyle.Star5},
                            {LineType, eShapeStyle.Line},
                            {"SHAPE_COMMENT", eShapeStyle.WedgeEllipseCallout},
                            {"SHAPE_ARROW", eShapeStyle.RightArrow}
                        };
        }

        public ShapePainter(ExcelWorksheet worksheet,
                            Dictionary<DateTime, int> columnsLookup,
                            Dictionary<int, RowDefinition> planRows)
        {
            _worksheet = worksheet;
            _columnsLookup = columnsLookup;
            _planRows = planRows;
        }

        public void DrawShape(Shape shapeObject)
        {
            try
            {
                DrawShapeUnsafe(shapeObject);
            }
            catch
            {
                // ignored
            }
        }

        private void DrawShapeUnsafe(Shape shapeObject)
        {
            if (!ShapesMap.TryGetValue(shapeObject.ShapeType, out var shapeType))
            {
                return;
            }

            var startColumnIndex = _columnsLookup[shapeObject.StartDate.Date] - 1;
            var endColumnIndex = _columnsLookup[shapeObject.EndDate.AddDays(-1).Date];

            var startRow = _planRows[shapeObject.RowStart];
            var endRow = _planRows[shapeObject.RowEnd];

            var startRowIndex = startRow.StartExcelRowIndex - 1;
            var endRowIndex = endRow.EndExcelRowIndex;

            if (shapeType == ShapesMap[LineType])
            {
                startRowIndex += (endRowIndex - startRowIndex) / 2;
                endRowIndex = startRowIndex;
            }

            var shape = _worksheet.Drawings.AddShape(shapeObject.ID.ToString(), shapeType);

            shape.From.Row = startRowIndex;
            shape.From.Column = startColumnIndex;
            shape.To.Row = endRowIndex;
            shape.To.Column = endColumnIndex;

            var appearance = AppearanceHelper.GetAppearance(shapeObject.Appearance);
            FormatShape(shape, appearance);
            SetRotation(_worksheet.Drawings.DrawingXml, appearance);
        }

        private static void FormatShape(ExcelShape shape, CellsAppearance appearance)
        {
            if (appearance.UseFillColor)
            {
                shape.Fill.Color = appearance.FillColor;
            }
            else
            {
                shape.Fill.Color = Color.Transparent;
                shape.Fill.Transparancy = 100;
            }

            shape.Border.Width = appearance.StrokeWidth;
            shape.Border.Fill.Color = appearance.StrokeColor;
        }

        private static void SetRotation(XmlDocument xml, CellsAppearance appearance)
        {
            if (appearance.RotationAngle == 0)
            {
                return;
            }

            var shapeNode = xml.LastChild.LastChild;
            var xdrNode = shapeNode.ChildNodes[2];
            var spPrNode = xdrNode.ChildNodes[1];

            var xfrmNode = xml.CreateNode(XmlNodeType.Element,
                                          "a:xfrm",
                                          "http://schemas.openxmlformats.org/drawingml/2006/main");
            xfrmNode = spPrNode.PrependChild(xfrmNode);

            var rotAttribute = xml.CreateAttribute("rot");
            rotAttribute.Value = (60000 * appearance.RotationAngle).ToString();
            xfrmNode.Attributes.Append(rotAttribute);
        }
    }
}