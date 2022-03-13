using System.Xml.Serialization;

namespace DXPlus.Charts;

/// <summary>
/// Specifies the possible positions for a legend.
/// </summary>
public enum ChartLegendPosition
{
    /// <summary>
    /// Legend is on the top
    /// </summary>
    [XmlAttribute("t")] Top,
    /// <summary>
    /// Legend is on the bottom
    /// </summary>
    [XmlAttribute("b")] Bottom,
    /// <summary>
    /// Legend is on the left
    /// </summary>
    [XmlAttribute("l")] Left,
    /// <summary>
    /// Legend is on the right
    /// </summary>
    [XmlAttribute("r")] Right,
    /// <summary>
    /// Legend is on the top/right
    /// </summary>
    [XmlAttribute("tr")] TopRight
}