using System.Drawing;
using System.Xml.Linq;

namespace DXPlus.Internal;

internal static class ShapeHelpers
{
    public static Color ParseColorElement(XElement element)
    {
        /*
         TODO: support other schemes
            <hslClr> (Hue, Saturation, Luminance Color Model)
            <prstClr> (Preset Color)
            <schemeClr> (Scheme Color)
            <scrgbClr> (RGB Color Model - Percentage Variant)
            <srgbClr> (RGB Color Model - Hex Variant)
            <sysClr> (System Color)
         */

        var rgb = element.Element(Namespace.DrawingMain + "srgbClr");
        if (rgb != null)
        {
            var value = rgb.GetValAttr();
            if (value != null)
            {
                return ColorTranslator.FromHtml($"#{value.Value}");
            }

            int.TryParse(rgb.AttributeValue("r") ?? "0", out int r);
            int.TryParse(rgb.AttributeValue("g") ?? "0", out int g);
            int.TryParse(rgb.AttributeValue("b") ?? "0", out int b);
            return Color.FromArgb(r, g, b);
        }

        return Color.Black;
    }
}