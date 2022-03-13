using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Border Types for tables, cells, and paragraphs
/// </summary>
internal enum BorderType
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

    /// <summary>
    /// Specifies the border between each paragraph in a set of paragraphs which have the same set of paragraph border settings.
    /// So if adjoining paragraphs have identical border settings, then there will be one border between them as specified
    /// by the between element. Otherwise the first paragraph will use its bottom border and the following paragraph will use its top border.
    /// </summary>
    Between,

    /// <summary>
    /// Specifies the border which may be displayed on the inside edge of the paragraph when the parent's
    /// section settings specify that the section shall be printed using mirrored margins.
    /// </summary>
    Bar,
}