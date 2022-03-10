using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Table Cell Border Types
/// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
/// </summary>
public enum TableCellBorderType
{
    /// <summary>
    /// No border
    /// </summary>
    None,

    /// <summary>
    /// Top edge
    /// </summary>
    Top,

    /// <summary>
    /// Bottom edge
    /// </summary>
    Bottom,

    /// <summary>
    /// Left edge
    /// </summary>
    Left,

    /// <summary>
    /// Right edge
    /// </summary>
    Right,

    /// <summary>
    /// Inside horizontal
    /// </summary>
    InsideH,

    /// <summary>
    /// Inside vertical
    /// </summary>
    InsideV,

    /// <summary>
    /// Top/Left to Bottom/Right
    /// </summary>
    [XmlAttribute("tl2br")] TopLeftToBottomRight,

    /// <summary>
    /// Top/Right to Bottom/Left
    /// </summary>
    [XmlAttribute("tr2bl")] TopRightToBottomLeft,
}