using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus.Charts;

/// <summary>
/// Axis base class
/// </summary>
public abstract class Axis : XElementWrapper
{
    private new XElement Xml => base.Xml!;

    /// <summary>
    /// ID of this Axis
    /// </summary>
    public string? Id => Xml!.Element(Namespace.Chart + "axId").GetVal();

    /// <summary>
    /// Return true if this axis is visible
    /// </summary>
    public bool IsVisible
    {
        get => Xml.Element(Namespace.Chart + "delete")?.GetVal() == "0";
        set => Xml.GetOrAddElement(Namespace.Chart + "delete").SetValue(value ? "1" : "0");
    }

    /// <summary>
    /// Axis position
    /// </summary>
    public ChartLegendPosition? Position
    {
        get => Xml.Element(Namespace.Chart + "axPos")?.GetVal().TryGetEnumValue<ChartLegendPosition>(out var pos) == true ? pos : null;
        set => Xml.GetOrAddElement(Namespace.Chart + "axPos").SetAttributeValue(Name.MainVal, value?.GetEnumName());
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="element">The XML this chart is represented by</param>
    protected Axis(XElement element)
    {
        base.Xml = element;
    }
}