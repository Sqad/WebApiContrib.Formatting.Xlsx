using System;
using System.Drawing;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Helpers
{
    internal class NeutralColorGenerator
    {
        private readonly Color _seedColor = Color.White;
        private readonly Random _random = new Random();

        public Color GetNextColor()
        {
            var red = _random.Next(256);
            var green = _random.Next(256);
            var blue = _random.Next(256);

            red = (red + _seedColor.R) / 2;
            green = (green + _seedColor.G) / 2;
            blue = (blue + _seedColor.B) / 2;

            return Color.FromArgb(red, green, blue);
        }
    }
}
