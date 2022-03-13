using System.Xml.Serialization;

namespace DXPlus.Charts;

/// <summary>
/// Specifies the possible directions for a bar chart.
/// </summary>
public enum BarDirection
{
    /// <summary>
    /// Vertical
    /// </summary>
    [XmlAttribute("col")] Column,

    /// <summary>
    /// Horizontal
    /// </summary>
    Bar
}