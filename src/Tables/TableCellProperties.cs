using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents the properties applied to a single table cell {tcPr}
/// </summary>
public sealed class TableCellProperties : XElementWrapper
{
    private new XElement Xml => base.Xml!;

    /// <summary>
    /// Gets or Sets the shading fill for this Cell.
    /// </summary>
    public Shading? Shading
    {
        get => Shading.FromElement(Xml);
        set => Shading.SetElementValue(Xml, value);
    }

    /// <summary>
    /// Get the applied gridSpan based on cell merges.
    /// </summary>
    public int GridSpan =>
        int.TryParse(Xml.Element(Namespace.Main + "gridSpan")?
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
    /// Get or set the text direction for the table cell
    /// </summary>
    public TextDirection? TextDirection
    {
        get => Xml.Element(Namespace.Main + "textDirection")?.GetVal()?.TryGetEnumValue(out TextDirection result) == true
            ? result
            : null;
        set => Xml.GetOrAddElement(Namespace.Main + "textDirection").SetAttributeValue(Name.MainVal, value?.GetEnumName());
    }

    /// <summary>
    /// Gets or sets the vertical alignment.
    /// </summary>
    public VerticalAlignment? VerticalAlignment
    {
        get => Xml.Element(Namespace.Main + "vAlign")?.GetVal()?.TryGetEnumValue(out VerticalAlignment result) == true
            ? result
            : null;
        set => Xml.GetOrAddElement(Namespace.Main + "vAlign").SetAttributeValue(Name.MainVal, value?.GetEnumName());
    }

    /// <summary>
    /// Specified width for the cell
    /// </summary>
    public TableElementWidth? CellWidth
    {
        get => new(Xml.Element(Namespace.Main + "tcW"));
        set
        {
            Xml.Element(Namespace.Main + "tcW")?.Remove();
            if (value == null
                || value.Type == null && value.Width == null) return;
            Xml.Add(new XElement(Namespace.Main + "tcW", value.Xml.Attributes()));
        }
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
        var marginAttribute = Xml.Element(Namespace.Main + "tcMar")?
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
            var margin = Xml.GetOrAddElement(Namespace.Main + "tcMar")
                .GetOrAddElement(Namespace.Main + name);

            margin.SetAttributeValue(Namespace.Main + "type", "dxa");
            margin.SetAttributeValue(Namespace.Main + "w", value);
        }
        else
        {
            Xml.Element(Namespace.Main + "tcMar")?.Element(Namespace.Main + name)?.Remove();
        }
    }

    /// <summary>
    /// Table border
    /// </summary>
    private static readonly XName tcBorders = Namespace.Main + "tcBorders";

    /// <summary>
    /// Top border for this table cell
    /// </summary>
    public Border? TopBorder
    {
        get => Border.FromElement(BorderType.Top, Xml, tcBorders);
        set => Border.SetElementValue(BorderType.Top, Xml, tcBorders, value);
    }

    /// <summary>
    /// Bottom border for this table cell
    /// </summary>
    public Border? BottomBorder
    {
        get => Border.FromElement(BorderType.Bottom, Xml, tcBorders);
        set => Border.SetElementValue(BorderType.Bottom, Xml, tcBorders, value);
    }

    /// <summary>
    /// Left border for this table cell
    /// </summary>
    public Border? LeftBorder
    {
        get => Border.FromElement(BorderType.Left, Xml, tcBorders);
        set => Border.SetElementValue(BorderType.Left, Xml, tcBorders, value);
    }

    /// <summary>
    /// Right border for this table cell
    /// </summary>
    public Border? RightBorder
    {
        get => Border.FromElement(BorderType.Right, Xml, tcBorders);
        set => Border.SetElementValue(BorderType.Right, Xml, tcBorders, value);
    }

    /// <summary>
    /// Inside horizontal border for this table cell
    /// </summary>
    public Border? InsideHorizontalBorder
    {
        get => Border.FromElement(BorderType.InsideH, Xml, tcBorders);
        set => Border.SetElementValue(BorderType.InsideH, Xml, tcBorders, value);
    }

    /// <summary>
    /// Inside horizontal border for this table cell
    /// </summary>
    public Border? InsideVerticalBorder
    {
        get => Border.FromElement(BorderType.InsideV, Xml, tcBorders);
        set => Border.SetElementValue(BorderType.InsideV, Xml, tcBorders, value);
    }

    /// <summary>
    /// Diagonal border for this table cell
    /// </summary>
    public Border? TopLeftToBottomRightDiagonalBorder
    {
        get => Border.FromElement(BorderType.TopLeftToBottomRight, Xml, tcBorders);
        set => Border.SetElementValue(BorderType.TopLeftToBottomRight, Xml, tcBorders, value);
    }

    /// <summary>
    /// Diagonal border for this table cell
    /// </summary>
    public Border? TopRightToBottomLeftDiagonalBorder
    {
        get => Border.FromElement(BorderType.TopRightToBottomLeft, Xml, tcBorders);
        set => Border.SetElementValue(BorderType.TopRightToBottomLeft, Xml, tcBorders, value);
    }

    /// <summary>
    /// Public constructor
    /// </summary>
    public TableCellProperties() : this(new XElement(Name.TableCellProperties))
    {
    }

    /// <summary>
    /// Formatting properties for the cell
    /// </summary>
    /// <param name="xml"></param>
    internal TableCellProperties(XElement xml)
    {
        base.Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        if (base.Xml.Name != Name.TableCellProperties)
            throw new ArgumentException($"Unexpected XML element {base.Xml.Name}, expected {Name.TableCellProperties}",
                nameof(xml));
    }
}