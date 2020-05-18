using OfficeOpenXml;
using OfficeOpenXml.Drawing;
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
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class PicturesPainter
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly Dictionary<DateTime, int> _columnsLookup;
        private readonly Dictionary<int, RowDefinition> _planRows;

        public PicturesPainter(ExcelWorksheet worksheet,
                               Dictionary<DateTime, int> columnsLookup,
                               Dictionary<int, RowDefinition> planRows)
        {
            _worksheet = worksheet;
            _columnsLookup = columnsLookup;
            _planRows = planRows;
        }

        public void DrawPictures(List<Picture> pictures, FormattedPlanViewMode viewMode)
        {
            if (!pictures.Any())
            {
                return;
            }

            var imagesLookup = DownloadAllImages(pictures);
            foreach (var picture in pictures)
            {
                try
                {
                    DrawPicture(picture, imagesLookup);
                }
                catch
                {
                    // ignored
                }
            }
        }

        private void DrawPicture(Picture picture, Dictionary<string, Image> imagesLookup)
        {
            var startColumnIndex = _columnsLookup[picture.StartDate.Date] - 1;
            var endColumnIndex = _columnsLookup[picture.EndDate.AddDays(-1).Date];

            var startRow = _planRows[picture.RowStart];
            var endRow = _planRows[picture.RowEnd];

            var startRowIndex = startRow.StartExcelRowIndex - 1;
            var endRowIndex = endRow.StartExcelRowIndex;

            if (!imagesLookup.TryGetValue(picture.ImageUrl, out var image))
            {
                return;
            }

            var shape = _worksheet.Drawings.AddPicture(picture.ID.ToString(), image);

            var appearance = AppearanceHelper.GetAppearance(picture.Appearance);
            CalculateSize(shape, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex, image, appearance);

            FormatPicture(shape, appearance);
            SetTransparency(_worksheet.Drawings.DrawingXml, appearance);
        }

        private void CalculateSize(ExcelPicture shape,
                                   int startRowIndex,
                                   int endRowIndex,
                                   int startColumnIndex,
                                   int endColumnIndex,
                                   Image image,
                                   CellsAppearance appearance)
        {
            if (appearance.UseImageFillSizing)
            {
                shape.From.Row = startRowIndex;
                shape.From.Column = startColumnIndex;
                shape.To.Row = endRowIndex;
                shape.To.Column = endColumnIndex;

                return;
            }

            shape.From.Row = startRowIndex;
            shape.From.Column = startColumnIndex;

            const double columnCoefficient = 5;

            var targetWidth = 0;
            var targetHeight = 0;
            for (var i = startColumnIndex; i < endColumnIndex; i++)
            {
                targetWidth += (int) (_worksheet.Column(i).Width * columnCoefficient);
            }

            for (var i = startRowIndex; i <= endRowIndex; i++)
            {
                targetHeight += (int) _worksheet.Row(i).Height;
            }

            var sourceWidth = image.Width;
            var sourceHeight = image.Height;

            var percentW = targetWidth / (float) sourceWidth;
            var percentH = targetHeight / (float) sourceHeight;

            var percent = percentH < percentW ? percentH : percentW;

            var newWidth = (int) (sourceWidth * percent);
            var newHeight = (int) (sourceHeight * percent);

            shape.SetSize(newWidth, newHeight);

            //note: temporary fix since it doesn't works with some images
            //shape.From.Row = startRowIndex;
            //shape.From.Column = startColumnIndex;

            //double newWidth = 0;
            //double newHeight = 0;
            //double coeff = 6.5;
            //for (int i = startColumnIndex; i <= endColumnIndex; i++)
            //{
            //    newWidth += _worksheet.Column(i).Width;
            //}
            //for (int i = startRowIndex; i <= endRowIndex; i++)
            //{
            //    newHeight += _worksheet.Row(i).Height;
            //}

            //newWidth *= coeff;

            //if ((newWidth > image.Width) || (newHeight > image.Height))
            //{
            //    if (newWidth > image.Width)
            //    {
            //        double widthOffset = (newWidth - image.Width) / 2.0;
            //        double currWidth = 0;
            //        int i;
            //        for (i = startColumnIndex; i <= endColumnIndex; i++)
            //        {
            //            if (widthOffset <= _worksheet.Column(i).Width * coeff + currWidth)
            //            {
            //                break;
            //            }
            //            currWidth += _worksheet.Column(i).Width * coeff;
            //        }
            //        shape.From.Column = i;

            //    }

            //    if (newHeight > image.Height)
            //    {
            //        double heightOffset = (newHeight - image.Height) / 2.0;
            //        double currHeight = 0;
            //        int i;
            //        for (i = startRowIndex; i <= endRowIndex; i++)
            //        {
            //            if (heightOffset <= _worksheet.Row(i).Height + currHeight)
            //            {
            //                break;
            //            }
            //            currHeight += _worksheet.Row(i).Height;
            //        }
            //        shape.From.Row = i;
            //    }
            //}

            //shape.SetSize(image.Width, image.Height);
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