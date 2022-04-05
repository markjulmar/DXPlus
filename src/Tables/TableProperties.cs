using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// This manages the [tblPr] element in a Word document structure which holds all table-level properties.
/// </summary>
public sealed class TableProperties : XElementWrapper
{
    private readonly Table? tableOwner;
    private new XElement Xml => base.Xml!;

    /// <summary>
    /// The conditional formatting applied to the table
    /// </summary>
    public TableConditionalFormatting ConditionalFormatting
    {
        get => ReadTableConditionalFormatting();
        set => WriteTableConditionalFormat(value);
    }

    /// <summary>
    /// Preferred width for the table
    /// </summary>
    public TableElementWidth? TableWidth
    {
        get => new(Xml.Element(Namespace.Main + "tblW"));
        set
        {
            Xml.Element(Namespace.Main + "tblW")?.Remove();
            if (value == null
                || value.Type == null && value.Width == null) return;
            Xml.Add(value.Xml);
        }
    }

    /// <summary>
    /// True if the table will auto-fit the contents. This corresponds to the {tblLayout.type} value of the table properties.
    /// </summary>
    public TableLayout? TableLayout
    {
        get => Xml.Element(Namespace.Main + "tblLayout")?
                    .AttributeValue(Namespace.Main + "type")?
                    .TryGetEnumValue<TableLayout>(out var result) == true ? result : null;

        set => Xml.GetOrAddElement(Namespace.Main + "tblLayout")
                    .SetAttributeValue(Namespace.Main + "type", value?.GetEnumName());
    }

    /// <summary>
    /// Specifies the alignment of the current table with respect to the text margins in the current section
    /// </summary>
    public Alignment? Alignment
    {
        get => Xml.Element(Name.Alignment)
            .GetVal().TryGetEnumValue(out Alignment result)
            ? result
            : null;

        set => Xml.GetOrAddElement(Name.Alignment)
            .SetAttributeValue(Name.MainVal, value?.GetEnumName());
    }

    /// <summary>
    /// The design\style to apply to this table.
    /// </summary>
    public string? Design
    {
        get => Xml.Element(Namespace.Main + "tblStyle")?.GetVal(null);
        set
        {
            var style = Xml.GetOrAddElement(Namespace.Main + "tblStyle");
            if (value == null)
            {
                style.Remove();
            }
            else
            {
                style.SetAttributeValue(Name.MainVal, value);
            }

            if (tableOwner?.InDocument == true)
                tableOwner.EnsureTableStyleInDocument();
        }
    }

    /// <summary>
    /// Indentation in dxa units
    /// </summary>
    public double? Indent
    {
        get
        {
            var value = Xml.Element(Name.TableIndent)?.Attribute(Namespace.Main + "w");
            if (value != null && double.TryParse(value.Value, out var indentUnits))
                return indentUnits;

            value?.Remove();
            return null;
        }
        set
        {
            XElement tblIndent = Xml.GetOrAddElement(Name.TableIndent);
            if (value is null or < 0)
            {
                tblIndent.Remove();
            }
            else
            {
                tblIndent.SetAttributeValue(Namespace.Main + "type", "dxa");
                tblIndent.SetAttributeValue(Namespace.Main + "w", value);
            }
        }
    }

    /// <summary>
    /// Table border
    /// </summary>
    private static readonly XName tblBorders = Namespace.Main + "tblBorders";

    /// <summary>
    /// Top border for this table
    /// </summary>
    public Border? TopBorder
    {
        get => Border.FromElement(BorderType.Top, Xml, tblBorders);
        set => Border.SetElementValue(BorderType.Top, Xml, tblBorders, value);
    }

    /// <summary>
    /// Bottom border for this table
    /// </summary>
    public Border? BottomBorder
    {
        get => Border.FromElement(BorderType.Bottom, Xml, tblBorders);
        set => Border.SetElementValue(BorderType.Bottom, Xml, tblBorders, value);
    }

    /// <summary>
    /// Left border for this table
    /// </summary>
    public Border? LeftBorder
    {
        get => Border.FromElement(BorderType.Left, Xml, tblBorders);
        set => Border.SetElementValue(BorderType.Left, Xml, tblBorders, value);
    }

    /// <summary>
    /// Right border for this table
    /// </summary>
    public Border? RightBorder
    {
        get => Border.FromElement(BorderType.Right, Xml, tblBorders);
        set => Border.SetElementValue(BorderType.Right, Xml, tblBorders, value);
    }

    /// <summary>
    /// Inside horizontal border for this table
    /// </summary>
    public Border? InsideHorizontalBorder
    {
        get => Border.FromElement(BorderType.InsideH, Xml, tblBorders);
        set => Border.SetElementValue(BorderType.InsideH, Xml, tblBorders, value);
    }

    /// <summary>
    /// Inside horizontal border for this table
    /// </summary>
    public Border? InsideVerticalBorder
    {
        get => Border.FromElement(BorderType.InsideV, Xml, tblBorders);
        set => Border.SetElementValue(BorderType.InsideV, Xml, tblBorders, value);
    }

    /// <summary>
    /// This determines how many rows constitute each of the row bands
    /// for the current table, allowing row band formatting to be applied to groups of rows
    /// (rather than just single alternating rows) when the table is formatted.
    /// </summary>
    public int? RowBands
    {
        get => int.TryParse(Xml.Element(Namespace.Main + "tblStyleRowBandSize")?.GetVal(), out var result) ? result : null;
        set => Xml.AddElementVal(Namespace.Main + "tblStyleRowBandSize", value);
    }

    /// <summary>
    /// This determines how many columns constitute each of the column bands
    /// for the current table, allowing column band formatting to be applied to groups of columns
    /// (rather than just single alternating cols) when the table is formatted.
    /// </summary>
    public int? ColumnBands
    {
        get => int.TryParse(Xml.Element(Namespace.Main + "tblStyleColBandSize")?.GetVal(), out var result) ? result : null;
        set => Xml.AddElementVal(Namespace.Main + "tblStyleColBandSize", value);
    }

    /// <summary>
    /// Gets or Sets the value of the Table Caption (Alternate Text Title) of this table.
    /// </summary>
    public string TableCaption
    {
        get => Xml.Element(Namespace.Main + "tblCaption")?.GetVal() ?? string.Empty;
        set
        {
            Xml.Descendants(Namespace.Main + "tblCaption").FirstOrDefault()?.Remove();
            Xml.Add(new XElement(Namespace.Main + "tblCaption", new XAttribute(Name.MainVal, value)));
        }
    }

    /// <summary>
    /// Gets or Sets the value of the Table Description (Alternate Text Description) of this table.
    /// </summary>
    public string TableDescription
    {
        get => Xml.Element(Namespace.Main + "tblDescription")?.GetVal() ?? string.Empty;
        set
        {
            Xml.Descendants(Namespace.Main + "tblDescription").FirstOrDefault()?.Remove();
            Xml.Add(new XElement(Namespace.Main + "tblDescription", new XAttribute(Name.MainVal, value)));
        }
    }

    /// <summary>
    /// Gets the table margin value in pixels for the specified margin
    /// </summary>
    /// <param name="type">Table margin type</param>
    /// <returns>The value for the specified margin in pixels, null if it's not set.</returns>
    public double? GetDefaultCellMargin(TableCellMarginType type)
    {
        return double.TryParse(Xml.Element(Namespace.Main + "tblCellMar")?
            .Element(Namespace.Main + type.GetEnumName())?
            .AttributeValue(Namespace.Main + "w"), out double result) ? result : null;
    }

    /// <summary>
    /// Set the specified cell margin for the table-level.
    /// </summary>
    /// <param name="type">The side of the cell margin.</param>
    /// <param name="margin">The value for the specified cell margin in dxa units.</param>
    public void SetDefaultCellMargin(TableCellMarginType type, double? margin)
    {
        if (margin != null)
        {
            var cellMargin = Xml.GetOrAddElement(Namespace.Main + "tblCellMar")
                .GetOrAddElement(Namespace.Main + type.GetEnumName());
            cellMargin.SetAttributeValue(Namespace.Main + "w", margin);
            cellMargin.SetAttributeValue(Namespace.Main + "type", "dxa");
        }
        else
        {
            var margins = Xml.Element(Namespace.Main + "tblCellMar");
            margins?.Element(Namespace.Main + type.GetEnumName())?.Remove();
            if (margins?.IsEmpty == true)
            {
                margins.Remove();
            }
        }
    }

    /// <summary>
    /// Formatting properties for a table
    /// </summary>
    public TableProperties() : this(CreateDefaultTableProperties(), null)
    {
    }

    /// <summary>
    /// Formatting properties for the paragraph
    /// </summary>
    /// <param name="xml"></param>
    /// <param name="owner">Table owner</param>
    internal TableProperties(XElement xml, Table? owner)
    {
        this.tableOwner = owner;
        base.Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        if (base.Xml.Name != Name.TableProperties)
            throw new ArgumentException($"Unexpected XML element {base.Xml.Name}, expected {Name.TableProperties}",
                nameof(xml));
    }

    /// <summary>
    /// Read the hex value for table conditional formatting and turn it back
    /// into an enumeration.
    /// </summary>
    /// <returns>Enum value</returns>
    private TableConditionalFormatting ReadTableConditionalFormatting()
    {
        string? value = Xml.Element(Namespace.Main + "tblLook")?.GetVal();
        if (!string.IsNullOrEmpty(value))
        {
            // It's represented as a hex value, we need to turn it back to
            // an integer base-10 value to use Enum.Parse.
            if (int.TryParse(value, System.Globalization.NumberStyles.HexNumber,
                    null, out int num))
            {
                return Enum.TryParse<TableConditionalFormatting>(
                    num.ToString(), out var tcf)
                    ? tcf
                    : TableConditionalFormatting.None;
            }
        }
        return TableConditionalFormatting.None;
    }

    /// <summary>
    /// Create a default "tblPr" element for a table.
    /// </summary>
    /// <returns></returns>
    internal static XElement CreateDefaultTableProperties()
    {
        return new XElement(Name.TableProperties,
            new XElement(Namespace.Main + "tblStyle",
                new XAttribute(Name.MainVal, TableDesign.Grid)),
            new XElement(Namespace.Main + "tblW",
                new XAttribute(Namespace.Main + "type", "auto"),
                new XAttribute(Namespace.Main + "w", 0)));
    }

    /// <summary>
    /// Write the element children for the TableConditionalFormatting
    /// </summary>
    /// <param name="format"></param>
    private void WriteTableConditionalFormat(TableConditionalFormatting format)
    {
        var e = Xml.GetOrAddElement(Namespace.Main + "tblLook");
        e.RemoveAttributes();

        e.Add(
            new XAttribute(Namespace.Main + "firstColumn", format.HasFlag(TableConditionalFormatting.FirstColumn) ? 1 : 0),
            new XAttribute(Namespace.Main + "lastColumn", format.HasFlag(TableConditionalFormatting.LastColumn) ? 1 : 0),
            new XAttribute(Namespace.Main + "firstRow", format.HasFlag(TableConditionalFormatting.FirstRow) ? 1 : 0),
            new XAttribute(Namespace.Main + "lastRow", format.HasFlag(TableConditionalFormatting.LastRow) ? 1 : 0),
            new XAttribute(Namespace.Main + "noHBand", format.HasFlag(TableConditionalFormatting.NoRowBand) ? 1 : 0),
            new XAttribute(Namespace.Main + "noVBand", format.HasFlag(TableConditionalFormatting.NoColumnBand) ? 1 : 0),
            new XAttribute(Name.MainVal, format.ToHex(4)));
    }
}