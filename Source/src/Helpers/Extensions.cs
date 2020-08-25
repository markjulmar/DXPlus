using System.Drawing;
using System.Globalization;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    internal static class Extensions
    {
        /// <summary>
        /// Converts a string to camelCase
        /// </summary>
        /// <param name="value">object to convert - will be converted to a string</param>
        /// <returns></returns>
        internal static string ToCamelCase(this object value)
        {
            string text = value?.ToString();
            if (!string.IsNullOrEmpty(text) && char.IsUpper(text[0]))
            {
                text = char.ToLower(text[0]) + text.Substring(1);
            }

            return text;
        }

        /// <summary>
        /// Convert a Color to it's #RGB hex value.
        /// </summary>
        /// <param name="color">Color to convert</param>
        /// <returns></returns>
        public static string ToHex(this Color color)
        {
            return $"{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        /// <summary>
        /// Convert the value of an attribute to a Color using ARGB
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public static Color ToColor(this XAttribute color)
        {
            if (color != null)
            {
                var rgb = int.Parse(color.Value.Replace("#", ""), NumberStyles.HexNumber) | 0xff000000;
                return Color.FromArgb((int) rgb);

            }

            return Color.Empty;
        }
    }
}
