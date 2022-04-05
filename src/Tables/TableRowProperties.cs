using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Row level properties
/// </summary>
public sealed class TableRowProperties : XElementWrapper
{
    private new XElement Xml => base.Xml!;

    /// <summary>
    /// Allow row to break across pages.
    /// The default value is true: Word will break the contents of the row across pages.
    /// If set to false, the contents of the row will not be split across pages, the
    /// entire row will be moved to the next page instead.
    /// </summary>
    public bool BreakAcrossPages
    {
        get => Xml.Element(Namespace.Main + "cantSplit") == null;
        set => Xml.SetElementValue(Namespace.Main + "cantSplit", value ? null : string.Empty);
    }

    /// <summary>
    /// Row height (can be null). Value is in dxa units
    /// </summary>
    public double? Height
    {
        get
        {
            var value = Xml.Element(Namespace.Main + "trHeight")?.GetValAttr();
            if (value == null || !double.TryParse(value.Value, out var heightInDxa))
            {
                value?.Remove();
                return null;
            }

            return heightInDxa;
        }
        set => SetHeight(value, true);
    }

    /// <summary>
    /// Specifies that the glyph representing the end character of current table row
    /// shall not be displayed in the current document.
    /// </summary>
    public bool Hidden
    {
        get => Xml.Element(Namespace.Main + "hidden")?.BoolAttributeValue(Name.MainVal) ?? false;
        set
        {
            if (value == false)
            {
                Xml.Element(Namespace.Main + "hidden")?.Remove();
            }
            else
            {
                Xml.GetOrAddElement(Namespace.Main + "hidden")
                    .SetAttributeValue(Name.MainVal, true);
            }
        }
    }

    /// <summary>
    /// Specifies the number of columns in the table before the first cell in the table row.
    /// </summary>
    public double? ColumnsBeforeFirstCell
    {
        get => double.TryParse(Xml.Element(Namespace.Main + "gridBefore")?.GetVal(), out var result)
            ? result
            : null;

        set
        {
            if (value == null)
            {
                Xml.Element(Namespace.Main + "gridBefore")?.Remove();
            }
            else
            {
                Xml.GetOrAddElement(Namespace.Main + "gridBefore")
                    .SetAttributeValue(Name.MainVal, value);
            }
        }
    }

    /// <summary>
    /// Specifies the number of columns in the table left after the last cell in the table row.
    /// </summary>
    public double? ColumnsAfterLastCell
    {
        get => double.TryParse(Xml.Element(Namespace.Main + "gridAfter")?.GetVal(), out var result)
                ? result
                : null;

        set
        {
            if (value == null)
            {
                Xml.Element(Namespace.Main + "gridAfter")?.Remove();
            }
            else
            {
                Xml.GetOrAddElement(Namespace.Main + "gridAfter")
                   .SetAttributeValue(Name.MainVal, value);
            }
        }
    }

    /// <summary>
    /// Specifies the preferred width for the total number of grid columns before this table row
    /// </summary>
    public TableElementWidth? PreferredWidthBefore
    {
        get => new(Xml.Element(Namespace.Main + "wBefore"));
        set
        {
            Xml.Element(Namespace.Main + "wBefore")?.Remove();
            if (value == null
                || value.Type == null && value.Width == null) return;
            Xml.Add(new XElement(Namespace.Main + "wBefore", value.Xml.Attributes()));
        }
    }

    /// <summary>
    /// Specifies the preferred width for the total number of grid columns after this table row
    /// </summary>
    public TableElementWidth? PreferredWidthAfter
    {
        get => new(Xml.Element(Namespace.Main + "wAfter"));
        set
        {
            Xml.Element(Namespace.Main + "wAfter")?.Remove();
            if (value == null
                || value.Type == null && value.Width == null) return;
            Xml.Add(new XElement(Namespace.Main + "wAfter", value.Xml.Attributes()));
        }
    }

    /// <summary>
    /// Set to true to make this row the table header row that will be repeated on each page
    /// </summary>
    public bool IsHeader
    {
        get => Xml.Element(Namespace.Main + "tblHeader") != null;
        set
        {
            var tblHeader = Xml.Element(Namespace.Main + "tblHeader");
            if (tblHeader == null && value)
            {
                Xml.SetElementValue(Namespace.Main + "tblHeader", string.Empty);
            }
            else if (tblHeader != null && !value)
            {
                tblHeader.Remove();
            }
        }
    }

    /// <summary>
    /// Gets or set this row's text alignment.
    /// </summary>
    public Alignment? Alignment
    {
        get => Xml.Element(Name.Alignment).GetVal()
            .TryGetEnumValue<Alignment>(out var result) ? result : null;
        set => Xml.AddElementVal(Name.Alignment, value?.GetEnumName());
    }

    /// <summary>
    /// Specifies the default table cell spacing (the spacing between adjacent cells and the edges of the table)
    /// for all cells in the parent row. If specified, this element specifies the minimum amount of space which
    /// shall be left between all cells in the table including the width of the table borders in the calculation.
    /// It is important to note that row-level cell spacing shall be added inside of the text margins, which shall
    /// be aligned with the innermost starting edge of the text extents in a cell without row-level indentation
    /// or cell spacing. Row-level cell spacing shall not increase the width of the overall table.
    /// </summary>
    public TableElementWidth? CellSpacing
    {
        get => new(Xml.Element(Namespace.Main + "tblCellSpacing"));
        set
        {
            Xml.Element(Namespace.Main + "tblCellSpacing")?.Remove();
            if (value == null
                || value.Type == null && value.Width == null) return;
            Xml.Add(new XElement(Namespace.Main + "tblCellSpacing", value.Xml.Attributes()));
        }
    }

    /// <summary>
    /// Helper method to set either the exact height or the min-height
    /// </summary>
    /// <param name="height">The height value to set in dxa units</param>
    /// <param name="exact">If true, the height will be forced, otherwise it will be treated as a minimum height, auto growing past it if need be.
    /// </param>
    private void SetHeight(double? height, bool exact)
    {
        if (height != null)
        {
            var trHeight = Xml.GetOrAddElement(Namespace.Main + "trHeight");
            trHeight.SetAttributeValue(Namespace.Main + "hRule", exact ? "exact" : "atLeast");
            trHeight.SetAttributeValue(Name.MainVal, height);
        }
        else
        {
            Xml.Element(Namespace.Main + "trHeight")?.Remove();
        }
    }

    /// <summary>
    /// Formatting properties for the row
    /// </summary>
    /// <param name="xml"></param>
    internal TableRowProperties(XElement xml)
    {
        base.Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        if (base.Xml.Name != Name.TableRowProperties)
            throw new ArgumentException($"Unexpected XML element {base.Xml.Name}, expected {Name.TableRowProperties}",
                nameof(xml));
    }
}