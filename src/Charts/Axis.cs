using System.Xml.Linq;

namespace DXPlus.Charts;

/// <summary>
/// Axis base class
/// </summary>
public abstract class Axis
{
    /// <summary>
    /// Axis xml element
    /// </summary>
    internal XElement Xml { get; }
    
    /// <summary>
    /// ID of this Axis
    /// </summary>
    public string? Id => Xml.Element(Namespace.Chart + "axId").GetVal();

    /// <summary>
    /// Return true if this axis is visible
    /// </summary>
    public bool IsVisible
    {
        get => Xml.Element(Namespace.Chart + "delete")?.GetVal() == "0";
        set => Xml.GetOrAddElement(Namespace.Chart + "delete").SetValue(value ? "1" : "0");
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="element">The XML this chart is represented by</param>
    protected Axis(XElement element)
    {
        Xml = element;
    }
}