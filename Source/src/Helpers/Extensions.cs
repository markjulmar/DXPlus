using System.Drawing;

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

    }
}
