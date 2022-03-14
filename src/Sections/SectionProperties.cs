using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// This manages a block of section properties (w:sectPr)
/// </summary>
public sealed class SectionProperties
{
    /// <summary>
    /// XML that makes up this element
    /// </summary>
    internal XElement Xml {get;}

    /// <summary>
    /// Revision id tied to this section
    /// </summary>
    public string? RevisionId
    {
        get => Xml.AttributeValue(Namespace.Main + "rsidR");
        set => Xml.SetAttributeValue(Namespace.Main + "rsidR", value);
    }

    /// <summary>
    /// Section break type
    /// </summary>
    public SectionBreakType Type
    {
        get => Xml.Element(Namespace.Main + "type").GetVal()
            .TryGetEnumValue<SectionBreakType>(out var result)
            ? result : SectionBreakType.NextPage;
        set => Xml.AddElementVal(Namespace.Main + "type", value.GetEnumName());
    }

    /// <summary>
    /// Indicates that Headers.First should be used on the first page.
    /// If this is FALSE, then Headers.First is not used in the doc.
    /// </summary>
    public bool DifferentFirstPage
    {
        get => Xml.Element(Namespace.Main + "titlePg") != null;

        set
        {
            var titlePg = Xml.Element(Namespace.Main + "titlePg");
            if (titlePg == null && value)
            {
                Xml.Add(new XElement(Namespace.Main + "titlePg", string.Empty));
            }
            else if (titlePg != null && !value)
            {
                titlePg.Remove();
            }
        }
    }

    /// <summary>
    /// Gets or Sets the Direction of content in this FirstParagraph.
    /// </summary>
    public Direction Direction
    {
        get => Xml.Element(Name.RightToLeft) == null ? Direction.LeftToRight : Direction.RightToLeft;

        set
        {
            if (value == Direction.RightToLeft)
            {
                Xml.GetOrAddElement(Name.RightToLeft);
            }
            else
            {
                Xml.Element(Name.RightToLeft)?.Remove();
            }
        }
    }

    /// <summary>
    /// Page width adjusted by margins.
    /// </summary>
    public double AdjustedPageWidth => PageWidth - LeftMargin - RightMargin;

    /// <summary>
    /// Page width value in dxa units
    /// </summary>
    public double PageWidth
    {
        get => double.TryParse(Xml.Element(Namespace.Main + "pgSz")?.AttributeValue(Namespace.Main + "w"), out var value) ? value : PageSize.LetterWidth;

        set => Xml.GetOrAddElement(Namespace.Main + "pgSz")
            .SetAttributeValue(Namespace.Main + "w", value);
    }

    /// <summary>
    /// Page height value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
    /// </summary>
    public double PageHeight
    {
        get => double.TryParse(Xml.Element(Namespace.Main + "pgSz")?.AttributeValue(Namespace.Main + "h"), out var value) ? value : PageSize.LetterHeight;
        set => Xml.GetOrAddElement(Namespace.Main + "pgSz").SetAttributeValue(Namespace.Main + "h", value);
    }

    /// <summary>
    /// Orientation
    /// </summary>
    public Orientation Orientation
    {
        get => Xml.AttributeValue(Namespace.Main + "pgSz", Namespace.Main + "orient")
            .TryGetEnumValue<Orientation>(out var result)
            ? result
            : Orientation.Portrait;

        set
        {
            if (value != Orientation)
            {
                double pw = PageWidth;
                double ph = PageHeight;

                var pgSz = Xml.GetOrAddElement(Namespace.Main + "pgSz");
                pgSz.SetAttributeValue(Namespace.Main + "orient", value.GetEnumName());
                pgSz.SetAttributeValue(Namespace.Main + "w", ph);
                pgSz.SetAttributeValue(Namespace.Main + "h", pw);
            }
        }
    }

    /// <summary>
    /// True to mirror the margins
    /// </summary>
    public bool MirrorMargins
    {
        get => Xml.Element(Namespace.Main + "mirrorMargins") != null;
        set
        {
            if (value)
            {
                Xml.SetElementValue(Namespace.Main + "mirrorMargins", string.Empty);
            }
            else
            {
                Xml.Element(Namespace.Main + "mirrorMargins")?.Remove();
            }
        }
    }

    /// <summary>
    /// Bottom margin value in dxa units.
    /// </summary>
    public double BottomMargin
    {
        get => GetMarginAttribute(Name.Bottom);
        set => SetMarginAttribute(Name.Bottom, value);
    }

    /// <summary>
    /// Left margin value in dxa units.
    /// </summary>
    public double LeftMargin
    {
        get => GetMarginAttribute(Name.Left);
        set => SetMarginAttribute(Name.Left, value);
    }

    /// <summary>
    /// Right margin value in dxa units.
    /// </summary>
    public double RightMargin
    {
        get => GetMarginAttribute(Name.Right);
        set => SetMarginAttribute(Name.Right, value);
    }

    /// <summary>
    /// Top margin value in dxa units.
    /// </summary>
    public double TopMargin
    {
        get => GetMarginAttribute(Name.Top);
        set => SetMarginAttribute(Name.Top, value);
    }

    /// <summary>
    /// Starting page number
    /// </summary>
    public int? StartPageNumber
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "pgNumType", Namespace.Main + "start"), out var result) ? result : null;
        set => Xml.GetOrAddElement(Namespace.Main + "pgNumType")
                .SetAttributeValue(Namespace.Main + "start", value?.ToString());
    }

    /// <summary>
    /// Gets or sets the vertical alignment.
    /// </summary>
    public VerticalAlignment? VerticalAlignment
    {
        get
        {
            var val = Xml.Element(Namespace.Main + "vAlign")?.GetValAttr();
            return val != null && val.TryGetEnumValue<VerticalAlignment>(out var result)
                ? result
                : null;
        }

        set
        {
            if (value == null)
            {
                Xml.Element(Namespace.Main + "vAlign")?.Remove();
            }
            else
            {
                Xml.AddElementVal(Namespace.Main + "vAlign", value.Value.GetEnumName());
            }
        }
    }

    // TODO: add pgBorders, textDirection

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="xml"></param>
    internal SectionProperties(XElement? xml)
    {
        Xml = xml ?? throw new ArgumentNullException(nameof(xml));
    }

    /// <summary>
    /// Get a margin
    /// </summary>
    /// <param name="name">Margin to get</param>
    /// <returns>Value in dxa units</returns>
    private double GetMarginAttribute(XName name)
    {
        var top = Xml.Element(Namespace.Main + "pgMar")?.Attribute(name);
        return top != null && double.TryParse(top.Value, out var value) ? value : 0;
    }

    /// <summary>
    /// Set a margin
    /// </summary>
    /// <param name="name">Margin to set</param>
    /// <param name="value">Value in dxa units</param>
    private void SetMarginAttribute(XName name, double value)
    {
        Xml.GetOrAddElement(Namespace.Main + "pgMar")
            .SetAttributeValue(name, value);
    }
}