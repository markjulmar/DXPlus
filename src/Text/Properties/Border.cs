using System.Globalization;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// This represents a border borderEdgeType along a side of an element.
/// </summary>
public class Border
{
    private readonly XElement element;

    /// <summary>
    /// Returns the color for this border borderEdgeType.
    /// </summary>
    public ColorValue Color
    {
        get => new(element.AttributeValue(Name.Color),
                element.AttributeValue(Name.ThemeColor), element.AttributeValue(Name.ThemeTint),
                element.AttributeValue(Name.ThemeShade));
        set => value.SetElementValues(element, Name.Color);
    }

    /// <summary>
    /// Specifies whether the border should be modified to create the appearance of a shadow. For right and bottom borders, this is done
    /// by duplicating the border below and right of the normal location. For the right and top borders, this is done by moving the
    /// border down and to the right of the original location.
    /// </summary>
    public bool? Shadow
    {
        get => bool.TryParse(element.AttributeValue(Name.Shadow), out var result) ? result : null;
        set => element.SetAttributeValue(Name.Shadow, value);
    }

    /// <summary>
    /// Specifies the spacing offset in dxa units
    /// </summary>
    public double? Spacing
    {
        // Represented in 20ths of a pt.
        get => double.TryParse(element.AttributeValue(Namespace.Main + "space"), out var result) ? result : null;
        set
        {
            if (value == null)
            {
                element.Attribute(Namespace.Main + "space")?.Remove();
            }
            else
            {
                element.AddElementVal(Namespace.Main + "space", value?.ToString(CultureInfo.InvariantCulture));
            }
        }
    }

    /// <summary>
    /// Specifies the width of the border. Paragraph borders are line borders, the width is specified in eighths of a point
    /// with a minimum value of two (1/4 of a point) and a maximum value of 96 (twelve points).
    /// </summary>
    public double? Size
    {
        get => double.TryParse(element.AttributeValue(Name.Size), out var result) ? result : null;
        set
        {
            if (value is < 2 or > 96)
                throw new ArgumentOutOfRangeException(nameof(Size));
            element.SetAttributeValue(Name.Size, value);
        }
    }

    /// <summary>
    /// Specifies the style of the border. Paragraph borders can be only line borders.
    /// </summary>
    public BorderStyle Style
    {
        get => Enum.TryParse<BorderStyle>(element.AttributeValue(Name.MainVal, "None"), ignoreCase:true, out var bd) ? bd : BorderStyle.None;
        set => element.SetAttributeValue(Name.MainVal, value.GetEnumName());
    }

    /// <summary>
    /// Border edge for an existing border edge.
    /// </summary>
    /// <param name="element"></param>
    internal Border(XElement element)
    {
        this.element = element;
    }
}