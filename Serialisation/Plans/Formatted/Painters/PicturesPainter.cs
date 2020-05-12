using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
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

        public int DrawPictures(List<Picture> pictures)
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
                shape.To.Row = endRowIndex;
                shape.To.Column = endColumnIndex;

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