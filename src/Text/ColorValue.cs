using System.Drawing;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// This wraps a color, tint, and shade property exposed by various objects.
/// </summary>
public sealed class ColorValue
{
    /// <summary>
    /// Specifies the color for this run.
    /// Value can be a specific color value (RGB), or Color.Empty to specify an automatic color.
    /// This value is mutually exclusive with the theme color values and is ignored if theme colors are specified.
    /// </summary>
    public Color? Color { get; }

    /// <summary>
    /// Specifies a theme color which should be used to color the element.
    /// The specified theme color is a reference to one of the predefined theme colors, located in the document's themes part,
    /// which allows for color information to be set centrally in the document.
    /// If this element is omitted, then no theme color is applied, and the Color property shall be used to determine the color.
    /// </summary>
    public ThemeColor? ThemeColor { get; }

    /// <summary>
    /// Specifies the tint value applied to the supplied theme color (if any) for this text.
    /// If a tint value is supplied, then it is applied to the RGB value of the color to determine the final color used.
    /// </summary>
    public byte? ThemeTint { get; }

    /// <summary>
    /// Specifies the shade value applied to the supplied theme color (if any) for this text.
    /// If the shade is supplied, then it is applied to the RGB value of the theme color to determine the final color used.
    /// </summary>
    public byte? ThemeShade { get; }

    /// <summary>
    /// Public constructor for single color
    /// </summary>
    /// <param name="value">Color to use</param>
    public ColorValue(Color value)
    {
        Color = value;
    }

    /// <summary>
    /// Converter from Color to ColorValue
    /// </summary>
    /// <param name="color"></param>
    public static implicit operator ColorValue(Color color) => new(color);

    /// <summary>
    /// Public constructor with theme color values
    /// </summary>
    /// <param name="color">Theme color enumeration</param>
    /// <param name="tint">Optional tint to apply</param>
    /// <param name="shading">Optional shading to apply</param>
    public ColorValue(ThemeColor color, byte? tint, byte? shading)
    {
        ThemeColor = color;
        ThemeTint = tint;
        ThemeShade = shading;
    }

    /// <summary>
    /// Internal constructor to set the object from an existing XML set of values.
    /// </summary>
    /// <param name="color">Color value</param>
    /// <param name="themeColor">Theme color</param>
    /// <param name="themeTint">Optional tint to apply</param>
    /// <param name="themeShading">Optional shading to apply</param>
    internal ColorValue(string? color, string? themeColor, string? themeTint, string? themeShading)
    {
        Color = color?.ToColor();
        ThemeColor = themeColor.TryGetEnumValue<ThemeColor>(out var result) ? result : null;
        ThemeTint = themeTint?.ToByte();
        ThemeShade = themeShading?.ToByte();
    }

    /// <summary>
    /// Constructor to build color values from a w:color element.
    /// </summary>
    /// <param name="element">Color element</param>
    internal ColorValue(XElement element) : this(element.GetVal(), element.AttributeValue(Name.ThemeColor),
            element.AttributeValue(Name.ThemeTint), element.AttributeValue(Name.ThemeShade))
    {
    }

    /// <summary>
    /// This returns whether the color value has any properties set.
    /// </summary>
    /// <returns>True if any values are set, false if the entire set of values is empty.</returns>
    internal bool IsEmpty() => Color == null && ThemeColor == null && ThemeTint == null && ThemeShade == null;

    /// <summary>
    /// Set the color values onto a specified XML element.
    /// </summary>
    /// <param name="element">XML element</param>
    /// <param name="mainVal">Name of main value</param>
    internal void SetElementValues(XElement element, XName? mainVal = null)
    {
        element.SetAttributeValue(mainVal ?? Name.MainVal, Color?.ToHex());
        element.SetAttributeValue(Name.ThemeColor, ThemeColor?.GetEnumName());
        element.SetAttributeValue(Name.ThemeTint, ThemeTint?.ToHex());
        element.SetAttributeValue(Name.ThemeShade, ThemeShade?.ToHex());
    }
}