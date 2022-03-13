using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// This wraps the TblW element found in table properties.
/// </summary>
public sealed class TableWidth
{
    internal const double PctMultiplier = 50.0;
    internal readonly XElement Xml;

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="element">TblW element or null to create one.</param>
    internal TableWidth(XElement? element)
    {
        this.Xml = element ?? new XElement(Namespace.Main + "tblW");
    }

    /// <summary>
    /// Constructor
    /// </summary>
    public TableWidth() : this(null)
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="units">Unit type</param>
    /// <param name="width">Width</param>
    public TableWidth(TableWidthUnit units, double width) : this(null)
    {
        Type = units;
        Width = width;
    }

    /// <summary>
    /// How the preferred table width is expressed.
    /// </summary>
    public TableWidthUnit? Type
    {
        get => Enum.TryParse<TableWidthUnit>(Xml.AttributeValue(Namespace.Main + "type"), 
                ignoreCase: true, out var tbw) ? tbw : null;

        set => Xml.SetAttributeValue(Namespace.Main + "type", value?.GetEnumName());
    }

    /// <summary>
    /// Preferred table width.
    /// </summary>
    public double? Width
    {
        get => double.TryParse(Xml.AttributeValue(Namespace.Main + "w"), out var d) ? d : null;

        set => Xml.SetAttributeValue(Namespace.Main + "w", value);
    }

    /// <summary>
    /// Create a table width from a percentage
    /// </summary>
    /// <param name="value">Percentage value</param>
    /// <returns>Table width</returns>
    public static TableWidth FromPercent(double value)
    {
        if (value is <= 0 or > 100) throw new ArgumentOutOfRangeException(nameof(value));
        return new TableWidth {Type = TableWidthUnit.Percentage, Width = value * PctMultiplier};
    }

    /// <summary>
    /// Create a table width from a Uom
    /// </summary>
    /// <param name="value">value</param>
    /// <returns></returns>
    public static TableWidth FromUom(Uom value)
    {
        if (value == null) throw new ArgumentNullException(nameof(value));
        return new TableWidth { Type = TableWidthUnit.Dxa, Width = value.Dxa };
    }

}