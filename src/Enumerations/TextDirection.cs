using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Specifies the text direction in a table cell
/// </summary>
public enum TextDirection
{
    /// <summary>
    /// Bottom to top, left to end
    /// </summary>
    [XmlAttribute("btLr")] BottomToTopLeftToEnd,

    /// <summary>
    /// Left to right, top to bottom
    /// </summary>
    [XmlAttribute("lrTb")] LeftToRightTopToBottom,

    /// <summary>
    /// Left to right, top to bottom vertical
    /// </summary>
    [XmlAttribute("lrTbV")] LeftToRightTopToBottomVertical,

    /// <summary>
    /// Top to bottom, left to right vertical
    /// </summary>
    [XmlAttribute("tbLrV")] TopToBottomLeftToRightVertical,

    /// <summary>
    /// Top to bottom, right to left
    /// </summary>
    [XmlAttribute("tbRl")] TopToBottomRightToLeft,

    /// <summary>
    /// Top to bottom, right to left vertical
    /// </summary>
    [XmlAttribute("tbRlV")] TopToBottomRightToLeftVertical,
};