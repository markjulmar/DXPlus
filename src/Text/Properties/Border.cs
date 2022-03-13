using System.Globalization;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// This represents a border along a side of an element.
/// </summary>
public sealed class Border : IEquatable<Border>
{
    private readonly BorderType type;
    private readonly XElement parent;
    private readonly XName? borderGroup;

    private XElement? Get(bool create)
    {
        if (!create)
        {
            return borderGroup != null 
                ? parent.Element(borderGroup)?.Element(Namespace.Main + type.GetEnumName()) 
                : parent.Element(Namespace.Main + type.GetEnumName());
        }
        return borderGroup != null 
            ? parent.GetOrAddElement(borderGroup, Namespace.Main + type.GetEnumName()) 
            : parent.GetOrAddElement(Namespace.Main + type.GetEnumName());
    }

    /// <summary>
    /// Set a border element onto the parent
    /// </summary>
    /// <param name="type">Border type</param>
    /// <param name="parent">Parent</param>
    /// <param name="borderGroup">Name of border group</param>
    /// <param name="value">Value to set, null to remove</param>
    internal static bool SetElementValue(BorderType type, XElement parent, XName? borderGroup, Border? value)
    {
        // Passed value is one of three things.
        // 1. New (unconnected) object: new Border(...)
        // 2. Object connected to a different element: x.Border = y.Border
        // 3. Same object: x.Border = x.Border

        XName tag = Namespace.Main + type.GetEnumName(); // top,left,bottom,right,etc.
        var existingTag = borderGroup != null ? parent.Element(borderGroup)?.Element(tag) : parent.Element(tag);
        var element = value?.Get(false);

        // Same object?
        if (existingTag == element) return false;

        // Remove the existing object and replace it.
        existingTag?.Remove();
        if (element == null || value?.IsEmpty()==true)
            return false;

        // Copy the border over.
        var owner = borderGroup != null ? parent.GetOrAddElement(borderGroup) : parent;
        owner.Add(new XElement(tag, element.Attributes()));
        return true;
    }

    /// <summary>
    /// Remove this element from the parent if the values do not impact the rendering.
    /// </summary>
    /// <returns></returns>
    private bool RemoveIfEmpty()
    {
        if (IsEmpty())
        {
            Get(false)?.Remove();
            return true;
        }

        return false;
    }

    /// <summary>
    /// True if this border has no values which will impact the rendering of the element.
    /// </summary>
    /// <returns></returns>
    private bool IsEmpty()
    {
        return Color.IsEmpty
            && Shadow == false
            && Frame == false
            && Style == BorderStyle.None
            && Spacing == null
            && Size == null;
    }

    /// <summary>
    /// Returns the color for this border borderEdgeType.
    /// </summary>
    public ColorValue Color
    {
        get
        {
            var element = Get(false);
            return new(element.AttributeValue(Name.Color),
                element.AttributeValue(Name.ThemeColor), element.AttributeValue(Name.ThemeTint),
                element.AttributeValue(Name.ThemeShade));
        }
        set
        {
            value.SetElementValues(Get(true)!, Name.Color);
            RemoveIfEmpty();
        }
    }

    /// <summary>
    /// Specifies whether the border should be modified to create the appearance of a shadow. For right and bottom borders, this is done
    /// by duplicating the border below and right of the normal location. For the right and top borders, this is done by moving the
    /// border down and to the right of the original location.
    /// </summary>
    public bool Shadow
    {
        get => bool.TryParse(Get(false).AttributeValue(Name.Shadow), out var result) && result;
        set
        {
            if (value == false)
            {
                Get(false)?.Attribute(Name.Shadow)?.Remove();
            }
            else
            {
                Get(true)!.SetAttributeValue(Name.Shadow, value.ToBoolean());
            }
            RemoveIfEmpty();
        }
    }

    /// <summary>
    /// Specifies the spacing offset in dxa units
    /// </summary>
    public double? Spacing
    {
        // Represented in 20ths of a pt.
        get => double.TryParse(Get(false).AttributeValue(Namespace.Main + "space"), out var result) ? result : null;
        set
        {
            if (value == null)
            {
                Get(false)?.Attribute(Namespace.Main + "space")?.Remove();
            }
            else
            {
                Get(true)!.AddElementVal(Namespace.Main + "space", value?.ToString(CultureInfo.InvariantCulture));
            }
            RemoveIfEmpty();
        }
    }

    /// <summary>
    /// Specifies the width of the border. Paragraph borders are line borders, the width is specified in eighths of a point
    /// with a minimum value of two (1/4 of a point) and a maximum value of 96 (twelve points).
    /// </summary>
    public double? Size
    {
        get => double.TryParse(Get(false).AttributeValue(Name.Size), out var result) ? result : null;
        set
        {
            if (value is < 2 or > 96)
                throw new ArgumentOutOfRangeException(nameof(Size));
            if (value == null)
            {
                Get(false)?.Attribute(Name.Size)?.Remove();
            }
            else
            {
                Get(true)!.SetAttributeValue(Name.Size, value);
            }
            RemoveIfEmpty();
        }
    }

    /// <summary>
    /// Specifies whether the specified border should be modified to create a frame effect by reversing the border's
    /// appearance from the edge nearest the text to the edge furthest from the text.
    /// If this attribute is omitted, then the border is not given any frame effect.
    /// </summary>
    public bool Frame
    {
        get => bool.TryParse(Get(false).AttributeValue(Namespace.Main + "frame"), out var result) && result;
        set
        {
            if (value == false)
            {
                Get(false)?.Attribute(Namespace.Main + "frame")?.Remove();
            }
            else
            {
                Get(true)!.SetAttributeValue(Namespace.Main + "frame", value.ToBoolean());
            }
            RemoveIfEmpty();
        }
    }

    /// <summary>
    /// Specifies the style of the border. Paragraph borders can be only line borders.
    /// </summary>
    public BorderStyle Style
    {
        get => Enum.TryParse<BorderStyle>(Get(false).AttributeValue(Name.MainVal, "None"), ignoreCase:true, out var bd) ? bd : BorderStyle.None;

        set
        {
            Get(true)!.SetAttributeValue(Name.MainVal, value.GetEnumName());
            RemoveIfEmpty();
        }
    }

    /// <summary>
    /// Returns a border object if the element exists on the given parent.
    /// </summary>
    /// <param name="type">Type</param>
    /// <param name="element">Parent element</param>
    /// <param name="borderGroup">Name of parent element</param>
    /// <returns>Border or null</returns>
    internal static Border? FromElement(BorderType type, XElement element, XName? borderGroup)
    {
        var border = new Border(type, element, borderGroup);
        return border.Get(false) != null ? border : null;
    }

    /// <summary>
    /// Border edge for an existing border edge.
    /// </summary>
    /// <param name="type">Border type</param>
    /// <param name="element">Parent element (properties)</param>
    /// <param name="borderGroup">Name of the border group (pBdr, tBdr, tcBdr, etc.)</param>
    private Border(BorderType type, XElement element, XName? borderGroup)
    {
        this.type = type;
        this.parent = element;
        this.borderGroup = borderGroup;
    }

    /// <summary>
    /// Create a Border edge for a table, paragraph or cell.
    /// </summary>
    public Border()
    {
        this.type = BorderType.None;
        this.parent = new XElement("prop", new XElement("bg"));
        this.borderGroup = "bg";
    }

    /// <summary>
    /// Constructor with parameters
    /// </summary>
    /// <param name="style"></param>
    /// <param name="size"></param>
    public Border(BorderStyle style, Uom size) : this()
    {
        Style = style;
        Size = size.Points * 8.0; //1/8ths of a point.
    }

    /// <summary>
    /// Determines equality between border objects
    /// </summary>
    /// <param name="other">Other object</param>
    /// <returns></returns>
    public bool Equals(Border? other) =>
        other is not null && (ReferenceEquals(this, other) ||
                              type == other.type && parent.Equals(other.parent) &&
                              borderGroup?.Equals(other.borderGroup) == true);

    /// <summary>
    /// Equality override
    /// </summary>
    /// <param name="obj"></param>
    /// <returns></returns>
    public override bool Equals(object? obj) 
        => ReferenceEquals(this, obj) || obj is Border other && Equals(other);

    /// <summary>
    /// Hashcode override
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() 
        => HashCode.Combine((int) type, parent, borderGroup);
}