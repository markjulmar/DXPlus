using System.Drawing;

namespace DXPlus
{
    internal static class Extensions
    {
        internal static string ToCamelCase(this object value)
        {
            string text = value?.ToString();
            if (!string.IsNullOrEmpty(text) && char.IsUpper(text[0]))
                text = char.ToLower(text[0]) + text.Substring(1);
            return text;
        }

        public static string ToHex(this Color source) => $"{source.R:X2}{source.G:X2}{source.B:X2}";
    }
}
