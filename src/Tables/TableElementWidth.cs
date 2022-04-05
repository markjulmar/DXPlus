using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// This wraps width values used on table and table cell elements.
/// </summary>
public sealed class TableElementWidth : IEquatable<TableElementWidth>
{
    internal const double PctMultiplier = 50.0;
    internal readonly XElement Xml;

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="element">TblW element or null to create one.</param>
    internal TableElementWidth(XElement? element)
    {
        // Assume table width, will be replaced by tableCell if this
        // gets assigned to a cell.
        this.Xml = element ?? new XElement(Namespace.Main + "tblW");
    }

    /// <summary>
    /// Constructor used with static methods and conversion operators.
    /// </summary>
    private TableElementWidth() : this(null)
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="units">Unit type</param>
    /// <param name="width">Width</param>
    public TableElementWidth(TableWidthUnit units, double width) : this(null)
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

        private init => Xml.SetAttributeValue(Namespace.Main + "type", value?.GetEnumName());
    }

    /// <summary>
    /// Converter from a double to a TableWidth - always converts to DXA.
    /// </summary>
    /// <param name="d">Value</param>
    public static implicit operator TableElementWidth(double d) => FromDxa(d);

    /// <summary>
    /// Converter from a Uom to a TableWidth - always converts to DXA.
    /// </summary>
    /// <param name="value">Value</param>
    public static implicit operator TableElementWidth(Uom value) => FromDxa(value);

    /// <summary>
    /// Preferred table width.
    /// </summary>
    public double? Width
    {
        get => double.TryParse(Xml.AttributeValue(Namespace.Main + "w"), out var d) ? d : null;

        private init => Xml.SetAttributeValue(Namespace.Main + "w", value);
    }

    /// <summary>
    /// Create a table width from a percentage
    /// </summary>
    /// <param name="value">Percentage value</param>
    /// <returns>Table width</returns>
    public static TableElementWidth FromPercent(double value)
    {
        if (value is <= 0 or > 100) throw new ArgumentOutOfRangeException(nameof(value));
        return new TableElementWidth {Type = TableWidthUnit.Percentage, Width = value * PctMultiplier};
    }

    /// <summary>
    /// Create a table width from a Uom
    /// </summary>
    /// <param name="value">value</param>
    /// <returns></returns>
    public static TableElementWidth FromDxa(Uom value)
    {
        if (value == null) throw new ArgumentNullException(nameof(value));
        return new TableElementWidth { Type = TableWidthUnit.Dxa, Width = value.Dxa };
    }

    /// <summary>
    /// Equality for TableWidth
    /// </summary>
    /// <param name="other">Other value</param>
    /// <returns></returns>
    public bool Equals(TableElementWidth? other) 
        => other is not null && (ReferenceEquals(this, other) || Xml.Equals(other.Xml));

    /// <summary>
    /// Equality override
    /// </summary>
    /// <param name="obj"></param>
    /// <returns></returns>
    public override bool Equals(object? obj) 
        => ReferenceEquals(this, obj) || obj is TableElementWidth other && Equals(other);

    /// <summary>
    /// Hashcode override
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();
}