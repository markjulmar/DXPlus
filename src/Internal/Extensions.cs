using System.Drawing;
using System.Globalization;

namespace DXPlus.Internal;

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
    /// Converts a boolean to a string
    /// </summary>
    /// <param name="value">Boolean value</param>
    /// <returns></returns>
    internal static string ToBoolean(this bool value) => value.ToString().ToLower();

    /// <summary>
    /// Remove any null values from an enumeration and cast to a non-null type.
    /// </summary>
    /// <typeparam name="T">Type</typeparam>
    /// <param name="values">Values</param>
    /// <returns>Values which are not null</returns>
    internal static IEnumerable<T> OmitNull<T>(this IEnumerable<T?> values)
    {
        foreach (var value in values)
            if (value != null) yield return value;
    }

    /// <summary>
    /// Convert a Color to it's #RGB hex value.
    /// </summary>
    /// <param name="color">Color to convert</param>
    /// <returns></returns>
    internal static string ToHex(this Color color)
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
    internal static string ToHex(this Enum value, int digits = -1)
    {
        int num = Convert.ToInt32(value);
        return digits <= 0 ? num.ToString("X") : num.ToString("X" + digits);
    }

    /// <summary>
    /// Convert a byte value to a hex string value
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    internal static string ToHex(this byte value) => value.ToString("2X");

    /// <summary>
    /// Convert a string value to a byte.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    internal static byte? ToByte(this string value) 
        => byte.TryParse(value, NumberStyles.HexNumber, null, out var b) ? b : null;

    /// <summary>
    /// Convert the value of a string to a System.Drawing.Color
    /// </summary>
    /// <param name="color">Hex color value RRGGBB</param>
    /// <returns>Color object</returns>
    internal static Color? ToColor(this string? color)
    {
        if (!string.IsNullOrEmpty(color))
        {
            if (string.Compare(color.Trim(), "auto", StringComparison.CurrentCultureIgnoreCase) == 0)
                return Color.Empty;

            if (color.StartsWith('#'))
                color = color[1..];
            if (uint.TryParse(color, NumberStyles.HexNumber, null, out var rgb))
                return Color.FromArgb((int)(rgb | 0xff000000));
        }
        return null;
    }
}