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
    /// Gets or Sets the shading fill for this Cell.
    /// </summary>
    public Shading Shading => new(tcPr);

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
    /// Specified width for the cell
    /// </summary>
    public TableWidth? CellWidth
    {
        get => new(tcPr.Element(Namespace.Main + "tcW"));
        set
        {
            tcPr.Element(Namespace.Main + "tcW")?.Remove();
            if (value == null
                || value.Type == null && value.Width == null) return;
            tcPr.Add(new XElement(Namespace.Main + "tcW", value.Xml.Attributes()));
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
    /// Table border
    /// </summary>
    private static readonly XName tcBorders = Namespace.Main + "tcBorders";

    /// <summary>
    /// Top border for this table cell
    /// </summary>
    public Border? TopBorder
    {
        get => new(BorderType.Top, tcPr, tcBorders);
        set => Border.SetElementValue(BorderType.Top, tcPr, tcBorders, value);
    }

    /// <summary>
    /// Bottom border for this table cell
    /// </summary>
    public Border? BottomBorder
    {
        get => new(BorderType.Bottom, tcPr, tcBorders);
        set => Border.SetElementValue(BorderType.Bottom, tcPr, tcBorders, value);
    }

    /// <summary>
    /// Left border for this table cell
    /// </summary>
    public Border? LeftBorder
    {
        get => new(BorderType.Left, tcPr, tcBorders);
        set => Border.SetElementValue(BorderType.Left, tcPr, tcBorders, value);
    }

    /// <summary>
    /// Right border for this table cell
    /// </summary>
    public Border? RightBorder
    {
        get => new(BorderType.Right, tcPr, tcBorders);
        set => Border.SetElementValue(BorderType.Right, tcPr, tcBorders, value);
    }

    /// <summary>
    /// Inside horizontal border for this table cell
    /// </summary>
    public Border? InsideHorizontalBorder
    {
        get => new(BorderType.InsideH, tcPr, tcBorders);
        set => Border.SetElementValue(BorderType.InsideH, tcPr, tcBorders, value);
    }

    /// <summary>
    /// Inside horizontal border for this table cell
    /// </summary>
    public Border? InsideVerticalBorder
    {
        get => new(BorderType.InsideV, tcPr, tcBorders);
        set => Border.SetElementValue(BorderType.InsideV, tcPr, tcBorders, value);
    }

    /// <summary>
    /// Diagonal border for this table cell
    /// </summary>
    public Border? TopLeftToBottomRightDiagonalBorder
    {
        get => new(BorderType.TopLeftToBottomRight, tcPr, tcBorders);
        set => Border.SetElementValue(BorderType.TopLeftToBottomRight, tcPr, tcBorders, value);
    }

    /// <summary>
    /// Diagonal border for this table cell
    /// </summary>
    public Border? TopRightToBottomLeftDiagonalBorder
    {
        get => new(BorderType.TopRightToBottomLeft, tcPr, tcBorders);
        set => Border.SetElementValue(BorderType.TopRightToBottomLeft, tcPr, tcBorders, value);
    }

    /// <summary>
    /// Determines equality for a table cell
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(TableCell? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a table cell
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as TableCell);

    /// <summary>
    /// Returns hashcode for this table
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();

}