using System.ComponentModel;
using System.Drawing;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// A single cell in a Word table. All content in a table is contained in a cell.
/// A cell also has several properties affecting its size, appearance, and how the content it contains is formatted.
/// </summary>
public sealed class TableCell : BlockContainer, IEquatable<TableCell>
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="row">Owner Row</param>
    /// <param name="xml">XML representing this cell</param>
    internal TableCell(TableRow row, XElement xml) : base(xml)
    {
        Row = row;
        if (Row.InDocument)
        {
            SetOwner(row.Document, row.PackagePart, false);
        }
    }

    /// <summary>
    /// Row owner
    /// </summary>
    public TableRow Row { get; }

    /// <summary>
    /// The table cell properties
    /// </summary>
    private XElement tcPr => Xml.GetOrAddElement(Namespace.Main + "tcPr");

    /// <summary>
    /// Gets or Sets the fill color of this Cell.
    /// TODO: use ColorValue
    /// </summary>
    public Color? FillColor
    {
        get => tcPr.Element(Namespace.Main + "shd")?.Attribute(Namespace.Main + "fill")?.ToColor();

        set
        {
            var shd = tcPr.GetOrAddElement(Namespace.Main + "shd");
            shd.SetAttributeValue(Name.MainVal, "clear");
            shd.SetAttributeValue(Name.Color, "auto");
            shd.SetAttributeValue(Namespace.Main + "fill", value?.ToHex());
        }
    }

    /// <summary>
    /// Get the applied gridSpan based on cell merges.
    /// </summary>
    public int GridSpan =>
        int.TryParse(tcPr.Element(Namespace.Main + "gridSpan")?
            .GetVal(), out int result)
            ? result
            : 1;

    /// <summary>
    /// Bottom margin in pixels.
    /// </summary>
    public double? BottomMargin
    {
        get => GetMargin(TableCellMarginType.Bottom.GetEnumName());
        set => SetMargin(TableCellMarginType.Bottom.GetEnumName(), value);
    }

    /// <summary>
    /// Left margin in pixels.
    /// </summary>
    public double? LeftMargin
    {
        get => GetMargin(TableCellMarginType.Left.GetEnumName());
        set => SetMargin(TableCellMarginType.Left.GetEnumName(), value);
    }

    /// <summary>
    /// Right margin in pixels.
    /// </summary>
    public double? RightMargin
    {
        get => GetMargin(TableCellMarginType.Right.GetEnumName());
        set => SetMargin(TableCellMarginType.Right.GetEnumName(), value);
    }

    /// <summary>
    /// Top margin in pixels.
    /// </summary>
    public double? TopMargin
    {
        get => GetMargin(TableCellMarginType.Top.GetEnumName());
        set => SetMargin(TableCellMarginType.Top.GetEnumName(), value);
    }

    /// <summary>
    /// Shading applied to the cell.
    /// TODO: use shading object
    /// </summary>
    public Color? Shading
    {
        get
        {
            var fill = tcPr.Element(Namespace.Main + "shd")?
                .Attribute(Namespace.Main + "fill");
            return fill == null ? null
                : ColorTranslator.FromHtml($"#{fill.Value}");
        }

        set
        {
            var shd = tcPr.GetOrAddElement(Namespace.Main + "shd");

            // The val attribute needs to be set to clear
            shd.SetAttributeValue(Name.MainVal, "clear");
            // The color attribute needs to be set to auto
            shd.SetAttributeValue(Name.Color, "auto");
            // The fill attribute needs to be set to the hex for this Color.
            shd.SetAttributeValue(Namespace.Main + "fill", value?.ToHex());
        }
    }

    /// <summary>
    /// Gets or sets all the text for a paragraph. This will replace any existing paragraph(s)
    /// tied to the table. The <seealso cref="Paragraph"/> property can also be used to manipulate
    /// the content of the cell.
    /// </summary>
    public string Text
    {
        get => string.Join('\n', Paragraphs.Select(p => p.Text).Where(p => !string.IsNullOrEmpty(p)));
        set
        {
            switch (Paragraphs.Count())
            {
                case 0:
                    this.AddParagraph(value);
                    break;
                case 1:
                    Paragraphs.First().SetText(value);
                    break;
                default:
                    Xml.Elements(Name.Paragraph).Remove();
                    this.AddParagraph(value);
                    break;
            }
        }
    }

    /// <summary>
    /// Get or set the text direction for the table cell
    /// </summary>
    public TextDirection? TextDirection
    {
        get => tcPr.Element(Namespace.Main + "textDirection")?.GetVal()?.TryGetEnumValue(out TextDirection result) == true
                ? result
                : null;
        set => tcPr.GetOrAddElement(Namespace.Main + "textDirection").SetAttributeValue(Name.MainVal, value?.GetEnumName());
    }

    /// <summary>
    /// Gets or sets the vertical alignment.
    /// </summary>
    public VerticalAlignment? VerticalAlignment
    {
        get => tcPr.Element(Namespace.Main + "vAlign")?.GetVal()?.TryGetEnumValue(out VerticalAlignment result) == true
            ? result
            : null;
        set => tcPr.GetOrAddElement(Namespace.Main + "vAlign").SetAttributeValue(Name.MainVal, value?.GetEnumName());
    }

    /// <summary>
    /// The units column widths for this cell are expressed in.
    /// </summary>
    public TableWidthUnit WidthUnit => Enum.TryParse<TableWidthUnit>(tcPr.Element(Namespace.Main + "tcW")?
        .AttributeValue(Namespace.Main + "type"), ignoreCase: true, out var tbw) ? tbw : TableWidthUnit.Auto;

    /// <summary>
    /// Width in WidthUnits
    /// </summary>
    public double? Width =>
        double.TryParse(
            tcPr.Element(Namespace.Main + "tcW")?.AttributeValue(Namespace.Main + "w"), out var value) 
            ? value : null;

    /// <summary>
    /// Sets the column width
    /// </summary>
    /// <param name="unitType">Unit type</param>
    /// <param name="value">Width in dxa or % units</param>
    public void SetWidth(TableWidthUnit unitType, double? value)
    {
        var tcW = tcPr.GetOrAddElement(Namespace.Main + "tcW");
        if (unitType == TableWidthUnit.None || value is null or < 0)
        {
            tcW.Remove();
            return;
        }

        tcW.SetAttributeValue(Namespace.Main + "type", unitType.GetEnumName());
        if (unitType == TableWidthUnit.Auto)
            value = 0;

        if (unitType == TableWidthUnit.Percentage)
            tcW.SetAttributeValue(Namespace.Main + "w", value.Value + "%");
        else
            tcW.SetAttributeValue(Namespace.Main + "w", value.Value);
    }

    /// <summary>
    /// Method to set all margins for the cell.
    /// </summary>
    /// <param name="margin"></param>
    public void SetMargins(double? margin)
    {
        LeftMargin = RightMargin = TopMargin = BottomMargin = margin;
    }

    /// <summary>
    /// Internal method to retrieve a specific margin by name
    /// </summary>
    /// <param name="name"></param>
    /// <returns>Margin in dxa units</returns>
    private double? GetMargin(string name)
    {
        var marginAttribute = tcPr.Element(Namespace.Main + "tcMar")?
                                  .Element(Namespace.Main + name)?.Attribute(Namespace.Main + "w");

        if (marginAttribute == null || !double.TryParse(marginAttribute.Value, out var margin))
        {
            marginAttribute?.Remove();
            return null;
        }

        return margin;
    }

    /// <summary>
    /// Internal method to set a specific margin by name
    /// </summary>
    /// <param name="name"></param>
    /// <param name="value">Value in dxa units"></param>
    private void SetMargin(string name, double? value)
    {
        if (value != null)
        {
            var margin = tcPr.GetOrAddElement(Namespace.Main + "tcMar")
                                    .GetOrAddElement(Namespace.Main + name);

            margin.SetAttributeValue(Namespace.Main + "type", "dxa");
            margin.SetAttributeValue(Namespace.Main + "w", value);
        }
        else
        {
            tcPr.Element(Namespace.Main + "tcMar")?.Element(Namespace.Main + name)?.Remove();
        }
    }

    /// <summary>
    /// Get a table cell border
    /// </summary>
    /// <param name="borderType">The table cell border to get</param>
    public Border? GetBorder(TableCellBorderType borderType)
    {
        var tcBorder = tcPr.Element(Namespace.Main + "tcBorders")?
                                   .Element(Namespace.Main + borderType.GetEnumName());
            
        return tcBorder != null ? new Border(tcBorder) : null;
    }

    /// <summary>
    /// Set outside borders to the given style
    /// </summary>
    public void SetOutsideBorders(BorderStyle style, Color color, double? spacing = 1, double size = 2)
    {
        SetBorder(TableCellBorderType.Top, style, color, spacing, size);
        SetBorder(TableCellBorderType.Left, style, color, spacing, size);
        SetBorder(TableCellBorderType.Right, style, color, spacing, size);
        SetBorder(TableCellBorderType.Bottom, style, color, spacing, size);
    }

    /// <summary>
    /// Set the table cell border
    /// </summary>
    public void SetBorder(TableCellBorderType borderType, BorderStyle style, Color color, double? spacing = 1, double size = 2)
    {
        if (size is < 2 or > 96)
            throw new ArgumentOutOfRangeException(nameof(Size));
        if (!Enum.IsDefined(typeof(ParagraphBorderType), borderType))
            throw new InvalidEnumArgumentException(nameof(borderType), (int)borderType, typeof(ParagraphBorderType));

        tcPr.Element(Namespace.Main + "tcBorders")?
            .Element(Namespace.Main.NamespaceName + borderType.GetEnumName())?.Remove();

        if (borderType == TableCellBorderType.None)
            return;

        // Set the border style
        var tcBorders = tcPr.GetOrAddElement(Namespace.Main + "tcBorders");
        var borderXml = new XElement(Namespace.Main + borderType.GetEnumName(),
            new XAttribute(Name.MainVal, style.GetEnumName()),
            new XAttribute(Name.Size, size));
        if (color != Color.Empty)
            borderXml.Add(new XAttribute(Name.Color, color.ToHex()));
        if (spacing != null)
            borderXml.Add(new XAttribute(Namespace.Main + "space", spacing));

        tcBorders.Add(borderXml);
    }

    /// <summary>
    /// Determines equality for a table cell
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(TableCell? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);
}