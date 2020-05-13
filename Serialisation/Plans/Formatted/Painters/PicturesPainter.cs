using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using SQAD.MTNext.Business.Models.FlowChart.Enums;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class PicturesPainter
    {
        private const int RowMultiplier = 3;

        private readonly ExcelWorksheet _worksheet;
        private readonly int _rowsOffset;
        private readonly Dictionary<DateTime, int> _columnsLookup;

        public PicturesPainter(ExcelWorksheet worksheet, int rowsOffset, Dictionary<DateTime, int> columnsLookup)
        {
            _worksheet = worksheet;
            _rowsOffset = rowsOffset;
            _columnsLookup = columnsLookup;
        }

        public int DrawPictures(List<Picture> pictures, FormattedPlanViewMode viewMode)
        {
            if (!pictures.Any())
            {
                return 0;
            }

            var imagesLookup = DownloadAllImages(pictures);

            var maxRowIndex = 0;
            foreach (var picture in pictures)
            {
                var startColumnIndex = _columnsLookup[picture.StartDate.Date] - 1;
                var endColumnIndex = _columnsLookup[picture.EndDate.AddDays(-1).Date];

                var startRowIndex = picture.RowStart * RowMultiplier + _rowsOffset - 3;
                var endRowIndex = picture.RowEnd * RowMultiplier + _rowsOffset;

                if (!imagesLookup.TryGetValue(picture.ImageUrl, out var image))
                {
                    continue;
                }

                var shape = _worksheet.Drawings.AddPicture(picture.ID.ToString(), image);

                shape.From.Row = startRowIndex;
                shape.From.Column = startColumnIndex;

                double newWidth = 0;
                double newHeight = 0;
                double coeff = 6.5;
                for (int i = startColumnIndex; i <= endColumnIndex; i++)
                {
                    newWidth += _worksheet.Column(i).Width;
                }
                for (int i = startRowIndex; i <= endRowIndex; i++)
                {
                    newHeight += _worksheet.Row(i).Height;
                }

                newWidth *= coeff;
                
                if ((newWidth > image.Width) || (newHeight > image.Height))
                {
                    if (newWidth > image.Width)
                    {
                        double widthOffset = (newWidth - image.Width) / 2.0;
                        double currWidth = 0;
                        int i;
                        for (i = startColumnIndex; i <= endColumnIndex; i++)
                        {
                            if (widthOffset <= _worksheet.Column(i).Width*coeff + currWidth)
                            {
                                break;
                            }
                            currWidth += _worksheet.Column(i).Width*coeff;
                        }
                        shape.From.Column = i;

                    }

                    if (newHeight > image.Height)
                    {
                        double heightOffset = (newHeight - image.Height) / 2.0;
                        double currHeight = 0;
                        int i;
                        for (i = startRowIndex; i <= endRowIndex; i++)
                        {
                            if (heightOffset <= _worksheet.Row(i).Height + currHeight)
                            {
                                break;
                            }
                            currHeight += _worksheet.Row(i).Height;
                        }
                        shape.From.Row = i;
                    }
                }

                shape.SetSize(image.Width, image.Height);
                

                var appearance = AppearanceHelper.GetAppearance(picture.Appearance);
                FormatPicture(shape, appearance);
                SetTransparency(_worksheet.Drawings.DrawingXml, appearance);

                if (endRowIndex > maxRowIndex)
                {
                    maxRowIndex = endRowIndex;
                }
            }

            return maxRowIndex;
        }

        private static void FormatPicture(ExcelPicture shape, CellsAppearance appearance)
        {
            if (!appearance.UseOutlineColor)
            {
                return;
            }

            shape.Border.LineStyle = eLineStyle.Solid;
            shape.Border.Fill.Color = appearance.OutlineColor;
        }

        private static void SetTransparency(XmlDocument xml, CellsAppearance appearance)
        {
            if (Math.Abs(appearance.Transparency - 1) < 0.001)
            {
                return;
            }

            var shapeNode = xml.LastChild.LastChild;
            var xdrNode = shapeNode.ChildNodes[2];
            var blipFillNode = xdrNode.ChildNodes[1];
            var blipNode = blipFillNode.ChildNodes[0];

            var alphaModFixNode = xml.CreateNode(XmlNodeType.Element, 
                                                 "a:alphaModFix",
                                                 "http://schemas.openxmlformats.org/drawingml/2006/main");
            alphaModFixNode = blipNode.AppendChild(alphaModFixNode);

            var amtAttribute = xml.CreateAttribute("amt");
            amtAttribute.Value = (100000 * appearance.Transparency).ToString();
            alphaModFixNode.Attributes.Append(amtAttribute);
        }

        private static Dictionary<string, Image> DownloadAllImages(IEnumerable<Picture> pictures)
        {
            var tasks = pictures.Select(x => x.ImageUrl)
                                .Distinct()
                                .Select(DownloadImage);

            var result = Task.WhenAll(tasks).Result;
            return result.ToDictionary(x => x.Item1, x => x.Item2);
        }

        private static async Task<(string, Image)> DownloadImage(string url)
        {
            using (var httpClient = new HttpClient())
            {
                using (var stream = await httpClient.GetStreamAsync(url))
                {
                    return (url, Image.FromStream(stream));
                }
            }
        }
    }
}