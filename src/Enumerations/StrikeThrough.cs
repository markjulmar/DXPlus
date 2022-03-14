using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Strike through styles
/// </summary>
public enum Strikethrough
{
    /// <summary>
    /// None
    /// </summary>
    None,

    /// <summary>
    /// Single-line strike through
    /// </summary>
    Strike,

    /// <summary>
    /// Double-line strike through
    /// </summary>
    [XmlAttribute("dstrike")] DoubleStrike
};