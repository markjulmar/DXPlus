using System.ComponentModel;
using System.Drawing;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus;

/// <summary>
/// This manages the [pPr] element in a Word document structure which holds all paragraph-level properties.
/// </summary>
public sealed class ParagraphProperties
{
    private const string DefaultStyle = "Normal";

    internal XElement Xml { get; }

    /// <summary>
    /// True to keep with the next element on the page.
    /// </summary>
    public bool KeepWithNext
    {
        get => Xml.Element(Name.KeepNext) != null;
        set => Xml.SetElementValue(Name.KeepNext, value ? string.Empty : null);
    }

    /// <summary>
    /// Keep all lines in this paragraph together on a page
    /// </summary>
    public bool KeepLinesTogether
    {
        get => Xml.Element(Name.KeepLines) != null;
        set => Xml.SetElementValue(Name.KeepLines, value ? string.Empty : null);
    }

    /// <summary>
    /// Gets or set this Paragraphs text alignment.
    /// </summary>
    public Alignment Alignment
    {
        get => Xml.Element(Name.ParagraphAlignment).GetVal()
            .TryGetEnumValue<Alignment>(out var result) ? result : Alignment.Left;
        set => Xml.AddElementVal(Name.ParagraphAlignment, value == Alignment.Left ? null : value.GetEnumName());
    }

    ///<summary>
    /// The style name of the paragraph.
    ///</summary>
    public string StyleName
    {
        get
        {
            var styleElement = Xml.Element(Name.ParagraphStyle);
            var attr = styleElement?.Attribute(Name.MainVal);
            return attr != null && !string.IsNullOrEmpty(attr.Value) ? attr.Value : DefaultStyle;
        }
        set
        {
            //TODO: should this be case-sensitive?
            if (string.IsNullOrWhiteSpace(value))
                value = DefaultStyle;
            Xml.AddElementVal(Name.ParagraphStyle, value != DefaultStyle ? value : null);
        }
    }

    /// <summary>
    /// Set the spacing between lines in this paragraph, in DXA units.
    /// </summary>
    public double? LineSpacing
    {
        get => GetLineSpacing("line");
        set => SetLineSpacing("line", value);
    }

    /// <summary>
    /// Specifies the logic which shall be used to calculate the line spacing of the associated paragraph or style.
    /// </summary>
    public LineRule? LineRule
    {
        get => Xml.Element(Name.Spacing).AttributeValue(Namespace.Main + "lineRule")
            .TryGetEnumValue<LineRule>(out var result) ? result : null;
        set => Xml.GetOrAddElement(Name.Spacing).SetAttributeValue(Namespace.Main + "lineRule", value?.GetEnumName());
    }

    /// <summary>
    /// Set the spacing between lines in this paragraph, in DXA units
    /// </summary>
    public double? LineSpacingAfter
    {
        get => GetLineSpacing("after");
        set => SetLineSpacing("after", value);
    }

    /// <summary>
    /// Set the spacing between lines in this paragraph, in DXA units
    /// </summary>
    public double? LineSpacingBefore
    {
        get => GetLineSpacing("before");
        set => SetLineSpacing("before", value);
    }

    /// <summary>
    /// Helper method to get spacing/xyz
    /// </summary>
    /// <param name="type">type of line spacing to retrieve</param>
    /// <returns>Value in dxa units, or null if not set</returns>
    private double? GetLineSpacing(string type)
    {
        var value = Xml.Element(Name.Spacing).AttributeValue(Namespace.Main + type, null);
        return value != null ? double.Parse(value) : null;
    }

    /// <summary>
    /// Helper method to set spacing/xyz
    /// </summary>
    /// <param name="type">type of line spacing to adjust</param>
    /// <param name="value">New value</param>
    private void SetLineSpacing(string type, double? value)
    {
        if (value != null)
        {
            Xml.GetOrAddElement(Name.Spacing).SetAttributeValue(Namespace.Main + type, value);
        }
        else
        {
            // Remove the 'spacing' element if it only specifies line spacing
            var el = Xml.Element(Name.Spacing);
            el?.Attribute(Namespace.Main + type)?.Remove();
            if (el?.HasAttributes == false)
            {
                el.Remove();
            }
        }
    }

    /// <summary>
    /// Set the left indentation in 1/20th pt for this FirstParagraph.
    /// </summary>
    public double? LeftIndent
    {
        get => double.TryParse(Xml.AttributeValue(Name.Indent, Name.Left), out var v) ? v : null;
        set
        {
            if (value == null)
            {
                var e = Xml.Element(Name.Indent);
                e?.Attribute(Name.Left)?.Remove();
                if (e?.HasAttributes == false)
                    e.Remove();
            }
            else
            {
                Xml.SetAttributeValue(Name.Indent, Name.Left, value);
            }
        }
    }

    /// <summary>
    /// Set the right indentation in 1/20th pt for this FirstParagraph.
    /// </summary>
    public double? RightIndent
    {
        get => double.TryParse(Xml.AttributeValue(Name.Indent, Name.Right), out var v) ? v : null;
        set
        {
            if (value == null)
            {
                var e = Xml.Element(Name.Indent);
                e?.Attribute(Name.Right)?.Remove();
                if (e?.HasAttributes == false)
                    e.Remove();
            }
            else
            {
                Xml.SetAttributeValue(Name.Indent, Name.Right, value);
            }
        }
    }

    /// <summary>
    /// The shade pattern applied to this paragraph
    /// TODO: move to Shade type.
    /// </summary>
    public ShadePattern? ShadePattern
    {
        get => Enum.TryParse<ShadePattern>(Xml.Element(Namespace.Main + "shd")?.GetVal(), ignoreCase: true, out var sp) ? sp : null;

        set
        {
            var e = Xml.Element(Namespace.Main + "shd");
            if (value == null)
            {
                if (e == null) return;
                value = DXPlus.ShadePattern.Clear;
            }
            e ??= HelperFunctions.CreateDefaultShadeElement(Xml);
            e.SetAttributeValue(Name.MainVal, value.Value.GetEnumName());
        }
    }

    /// <summary>
    /// Shade color used with pattern - use Color.Empty for "auto"
    /// TODO: move to Shade type.
    /// TODO: change to ColorValue
    /// </summary>
    public Color? ShadeColor
    {
        get
        {
            var color = Xml.Element(Namespace.Main + "shd")?.AttributeValue(Name.Color);
            return string.IsNullOrEmpty(color) ? null :
                color.ToLower() == "auto" ? Color.Empty : ColorTranslator.FromHtml($"#{color}");
        }

        set
        {
            var e = Xml.Element(Namespace.Main + "shd");
            if (value == null)
            {
                if (e == null) return;
                value = Color.Empty;
            }

            e ??= HelperFunctions.CreateDefaultShadeElement(Xml);
            e.SetAttributeValue(Name.Color, value == Color.Empty ? "auto" : value.Value.ToHex());
        }
    }

    /// <summary>
    /// Shade fill - use Color.Empty for "auto"
    /// TODO: move to Shade type.
    /// TODO: change to ColorValue
    /// </summary>
    public Color? ShadeFill
    {
        get
        {
            var color = Xml.Element(Namespace.Main + "shd")?.AttributeValue(Namespace.Main + "fill");
            return string.IsNullOrEmpty(color) ? null :
                color.ToLower() == "auto" ? Color.Empty : ColorTranslator.FromHtml($"#{color}");
        }

        set
        {
            var e = Xml.Element(Namespace.Main + "shd");
            if (value == null)
            {
                if (e == null) return;
                value = Color.Empty;
            }

            e ??= HelperFunctions.CreateDefaultShadeElement(Xml);
            e.SetAttributeValue(Namespace.Main + "fill", value == Color.Empty ? "auto" : value.Value.ToHex());
        }
    }

    /// <summary>
    /// Paragraph border
    /// </summary>
    private XElement? pBdr => Xml.Element(Namespace.Main + "pBdr");

    /// <summary>
    /// Top border for this paragraph
    /// </summary>
    public Border? TopBorder
    {
        get
        {
            var e = pBdr?.Element(Namespace.Main + ParagraphBorderType.Top.GetEnumName());
            return e == null ? null : new Border(e);
        }
    }

    /// <summary>
    /// Bottom border for this paragraph
    /// </summary>
    public Border? BottomBorder
    {
        get
        {
            var e = pBdr?.Element(Namespace.Main + ParagraphBorderType.Bottom.GetEnumName());
            return e == null ? null : new Border(e);
        }
    }

    /// <summary>
    /// Left border for this paragraph
    /// </summary>
    public Border? LeftBorder
    {
        get
        {
            var e = pBdr?.Element(Namespace.Main + ParagraphBorderType.Left.GetEnumName());
            return e == null ? null : new Border(e);
        }
    }

    /// <summary>
    /// Right border for this paragraph
    /// </summary>
    public Border? RightBorder
    {
        get
        {
            var e = pBdr?.Element(Namespace.Main + ParagraphBorderType.Right.GetEnumName());
            return e == null ? null : new Border(e);
        }
    }

    /// <summary>
    /// Between border for this paragraph
    /// </summary>
    public Border? BetweenBorder
    {
        get
        {
            var e = pBdr?.Element(Namespace.Main + ParagraphBorderType.Between.GetEnumName());
            return e == null ? null : new Border(e);
        }
    }

    /// <summary>
    /// Set all outside edges for the border
    /// </summary>
    public void SetBorders(BorderStyle style, Color color, double? spacing = 1, double size = 2, bool shadow = false)
    {
        if (size is < 2 or > 96)
            throw new ArgumentOutOfRangeException(nameof(size));

        SetBorder(ParagraphBorderType.Left, style, color, spacing, size, shadow);
        SetBorder(ParagraphBorderType.Top, style, color, spacing, size, shadow);
        SetBorder(ParagraphBorderType.Right, style, color, spacing, size, shadow);
        SetBorder(ParagraphBorderType.Bottom, style, color, spacing, size, shadow);
    }

    /// <summary>
    /// Set a specific border edge.
    /// </summary>
    /// <exception cref="InvalidEnumArgumentException"></exception>
    public void SetBorder(ParagraphBorderType borderType, BorderStyle style, Color color, double? spacing = 1, double size = 2, bool shadow = false)
    {
        if (size is < 2 or > 96)
            throw new ArgumentOutOfRangeException(nameof(Size));

        if (!Enum.IsDefined(typeof(ParagraphBorderType), borderType))
            throw new InvalidEnumArgumentException(nameof(borderType), (int)borderType, typeof(ParagraphBorderType));

        Xml.Element(Namespace.Main + "pBdr")?
            .Element(Namespace.Main + borderType.GetEnumName())?.Remove();

        if (style == BorderStyle.None)
            return;

        var pBdr = Xml.GetOrAddElement(Namespace.Main + "pBdr");
        var borderXml = new XElement(Namespace.Main + borderType.GetEnumName(),
            new XAttribute(Name.MainVal, style.GetEnumName()),
            new XAttribute(Name.Size, size));
        if (color != Color.Empty)
            borderXml.Add(new XAttribute(Name.Color, color.ToHex()));
        if (shadow)
            borderXml.Add(new XAttribute(Name.Shadow, true));
        if (spacing != null)
            borderXml.Add(new XAttribute(Namespace.Main + "space", spacing));

        pBdr.Add(borderXml);
    }

    /// <summary>
    /// Get or set the indentation of the first line of this FirstParagraph.
    /// </summary>
    public double? FirstLineIndent
    {
        get => double.TryParse(Xml.AttributeValue(Name.Indent, Name.FirstLine), out var v) ? v : null;
        set
        {
            var e = Xml.Element(Name.Indent);
            if (value == null)
            {
                e?.Attribute(Name.FirstLine)?.Remove();
                if (e?.HasAttributes == false)
                    e.Remove();
            }
            else
            {
                e?.Attribute(Name.Hanging)?.Remove();
                Xml.SetAttributeValue(Name.Indent, Name.FirstLine, value.Value);
            }
        }
    }

    /// <summary>
    /// Get or set the indentation of all but the first line of this FirstParagraph.
    /// </summary>
    public double? HangingIndent
    {
        get => double.TryParse(Xml.AttributeValue(Name.Indent, Name.Hanging), out var v) ? v : null;

        set
        {
            var e = Xml.Element(Name.Indent);
            if (value == null)
            {
                e?.Attribute(Name.Hanging)?.Remove();
                if (e?.HasAttributes == false)
                    e.Remove();
            }
            else
            {
                e?.Attribute(Name.FirstLine)?.Remove();
                Xml.SetAttributeValue(Name.Indent, Name.Hanging, value.Value);
            }
        }
    }

    /// <summary>
    /// Allow line breaking at the character level.
    /// </summary>
    public bool WordWrap
    {
        get => Xml.Element(Name.WordWrap)?.BoolAttributeValue(Name.MainVal) ?? false;
        set => Xml.GetOrAddElement(Name.WordWrap).SetAttributeValue(Name.MainVal, value ? 1 : 0);
    }

    /// <summary>
    /// Change the Right to Left paragraph layout.
    /// </summary>
    public Direction Direction
    {
        get => Xml.Element(Name.RightToLeft) == null ? Direction.LeftToRight : Direction.RightToLeft;
        set => Xml.SetElementValue(Name.RightToLeft, value == Direction.RightToLeft ? string.Empty : null);
    }

    /// <summary>
    /// Default formatting options for the paragraph
    /// </summary>
    public Formatting DefaultFormatting => new(Xml.CreateRunProperties());

    /// <summary>
    /// Formatting properties for the paragraph
    /// </summary>
    /// <param name="xml"></param>
    public ParagraphProperties(XElement? xml = null)
    {
        Xml = xml ?? new XElement(Name.ParagraphProperties);
    }
}