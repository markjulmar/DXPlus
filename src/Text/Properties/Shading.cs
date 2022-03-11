using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// Wrapper for all the shading properties for a run of text.
/// </summary>
public class Shading
{
    private readonly XElement properties;

    /// <summary>
    /// The Shade element
    /// </summary>
    private XElement? Shd(bool create)
    {
        return create 
            ? properties.GetOrAddElement(Namespace.Main + "shd")
            : properties.Element(Namespace.Main + "shd");
    }

    /// <summary>
    /// True if the parent has a shade element.
    /// </summary>
    public bool HasShading => Shd(false) != null;

    /// <summary>
    /// Specifies the pattern which shall be used to lay the pattern color over the background color for this paragraph shading.
    /// This pattern consists of a mask which is applied over the background shading color to get the locations where the pattern
    /// color should be shown based set enumeration values.
    /// </summary>
    public ShadePattern? Pattern
    {
        get => Shd(false).GetValAttr().TryGetEnumValue<ShadePattern>(out var result) ? result : null;
        set
        {
            if (value == null)
            {
                var shd = Shd(false);
                if (shd == null) return;
                shd.GetValAttr()?.Remove();
                if (!shd.HasAttributes)
                {
                    shd.Remove();
                    return;
                }
                
                value = ShadePattern.None;
            }

            Shd(true)!.SetAttributeValue(Name.MainVal, value.Value.GetEnumName());
        }
    }

    /// <summary>
    /// Specifies the color used for any foreground pattern for the pattern shading applied to this text.
    /// This color may either be presented as a specific color, or as Color.Empty to automatically determine the foreground shading color as appropriate.
    /// If the shading style specifies the use of no shading format or is omitted, then this property has no effect.
    /// Also, if the shading specifies the use of a theme color, then this value is superseded by the theme color value.
    /// If this attribute is omitted, then its value shall be assumed to be auto.
    /// </summary>
    public ColorValue? Color
    {
        get
        {
            var shd = Shd(false);
            return new(shd?.AttributeValue(Name.Color), shd?.AttributeValue(Name.ThemeColor), 
                shd?.AttributeValue(Name.ThemeTint), shd?.AttributeValue(Name.ThemeShade));
        }
        set
        {
            if (value == null || value.IsEmpty)
            {
                var shd = Shd(false);
                if (shd != null)
                {
                    shd.Attribute(Name.Color)?.Remove();
                    shd.Attribute(Name.ThemeColor)?.Remove();
                    shd.Attribute(Name.ThemeTint)?.Remove();
                    shd.Attribute(Name.ThemeShade)?.Remove();
                    if (!shd.HasAttributes) shd.Remove();
                }
                return;
            }

            value.SetElementValues(Shd(true)!, Name.Color);
        }
    }

    /// <summary>
    /// Specifies the color used for the background for this shading.
    /// This color may either be a specific color value or set to Empty to allow a consumer to automatically determine
    /// the background shading color as appropriate.
    /// If this attribute is omitted, then its value shall be assumed to be auto.
    /// </summary>
    public ColorValue? Fill
    {
        get
        {
            var shd = Shd(false);
            return new(shd?.AttributeValue(Name.Fill), shd?.AttributeValue(Name.ThemeFill),
                shd?.AttributeValue(Name.ThemeFillTint), shd?.AttributeValue(Name.ThemeFillShade));
        }
        set
        {
            var shd = Shd(false);
            if (value == null || value.IsEmpty)
            {
                if (shd != null)
                {
                    shd.Attribute(Name.Fill)?.Remove();
                    shd.Attribute(Name.ThemeFill)?.Remove();
                    shd.Attribute(Name.ThemeFillTint)?.Remove();
                    shd.Attribute(Name.ThemeFillShade)?.Remove();
                    if (!shd.HasAttributes) shd.Remove();
                }
                return;
            }

            shd ??= Shd(true);

            shd!.SetAttributeValue(Name.Fill, value.Color?.ToHex());
            shd.SetAttributeValue(Name.ThemeFill, value.ThemeColor?.GetEnumName());
            shd.SetAttributeValue(Name.ThemeFillTint, value.ThemeTint?.ToHex());
            shd.SetAttributeValue(Name.ThemeFillShade, value.ThemeShade?.ToHex());
        }
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="rp">Run properties</param>
    internal Shading(XElement rp)
    {
        properties = rp;
    }
}