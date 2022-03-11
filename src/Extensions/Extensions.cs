﻿using System;
using System.Drawing;
using System.Globalization;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Misc. .NET extensions to common types.
    /// </summary>
    internal static class Extensions
    {
        /// <summary>
        /// Converts a string to camelCase
        /// </summary>
        /// <param name="value">object to convert - will be converted to a string</param>
        /// <returns></returns>
        internal static string ToCamelCase(this object? value)
        {
            string? text = value?.ToString();
            if (!string.IsNullOrEmpty(text) && char.IsUpper(text[0]))
            {
                text = char.ToLower(text[0]) + text[1..];
            }

            return text ?? string.Empty;
        }

        /// <summary>
        /// Convert a Color to it's #RGB hex value.
        /// </summary>
        /// <param name="color">Color to convert</param>
        /// <returns></returns>
        public static string ToHex(this Color color)
        {
            return color == Color.Transparent || color == Color.Empty
                ? "auto" 
                : $"{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        /// <summary>
        /// Convert an enumeration value to hex.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="digits"></param>
        /// <returns></returns>
        public static string ToHex(this Enum value, int digits = -1)
        {
            int num = Convert.ToInt32(value);
            return digits <= 0 ? num.ToString("X") : num.ToString("X" + digits);
        }

        /// <summary>
        /// Convert a byte value to a hex string value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string ToHex(this byte value) => value.ToString("2X");

        /// <summary>
        /// Convert a string value to a byte.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static byte? ToByte(this string value) 
            => byte.TryParse(value, NumberStyles.HexNumber, null, out var b) ? b : null;

        /// <summary>
        /// Convert the value of a string to a System.Drawing.Color
        /// </summary>
        /// <param name="color">Hex color value RRGGBB</param>
        /// <returns>Color object</returns>
        public static Color? ToColor(this string? color)
        {
            if (color != null && string.Compare(color.Trim(), "auto", StringComparison.CurrentCultureIgnoreCase) != 0)
            {
                if (color.StartsWith('#'))
                    color = color[1..];
                if (uint.TryParse(color, NumberStyles.HexNumber, null, out var rgb))
                    return Color.FromArgb((int)(rgb | 0xff000000));
            }
            return null;
        }

        /// <summary>
        /// Convert the value of an attribute to a Color using ARGB
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public static Color ToColor(this XAttribute color)
        {
            if (color.Value.Trim().ToLower() == "auto")
                return Color.Transparent;

            var rgb = int.Parse(color.Value.Replace("#", ""), NumberStyles.HexNumber) | 0xff000000;
            return Color.FromArgb((int) rgb);
        }
    }
}
