using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Specifies the available types of custom tab stop, which determines the behavior of the
/// tab stop and the alignment which shall be applied to text entered at the current custom tab stop.
/// </summary>
public enum TabStopType
{
    /// <summary>
    /// Specifies that the current tab is a bar tab. A bar tab is a tab which does not result
    /// in a custom tab stop in the parent paragraph (this tab stop location shall be skipped
    /// when positioning custom tab characters), but instead shall be used to draw
    /// a vertical line (or bar) at this location in the parent paragraph.
    /// </summary>
    [XmlAttribute("bar")] Bar,

    /// <summary>
    /// Specifies that the current tab stop shall result in a location in the document where
    /// all following text is centered (i.e. all text runs following this tab stop and
    /// preceding the next tab stop shall be centered around the tab stop location).
    /// </summary>
    [XmlAttribute("center")] Centered,

    /// <summary>
    /// Specifies that the current tab stop is cleared and shall be removed and ignored
    /// when processing the contents of this document.
    /// </summary>
    [XmlAttribute("clear")] Clear,

    /// <summary>
    /// Specifies that the current tab stop shall result in a location in the document where
    /// all following text is aligned around the first decimal character in the following text runs.
    /// All text runs before the first decimal character shall be before the tab stop, all text runs
    /// after it shall be after the tab stop location
    /// </summary>
    [XmlAttribute("decimal")] Decimal,

    /// <summary>
    /// Specifies that the current tab stop shall result in a location in the document where
    /// all following text is left aligned (i.e. all text runs following this tab stop
    /// and preceding the next tab stop shall be left aligned with respect to the tab stop location).
    /// </summary>
    [XmlAttribute("left")] Left,

    /// <summary>
    /// Specifies that the current tab is a list tab, which is the tab stop between the numbering
    /// and the paragraph contents in a numbered paragraph.
    /// </summary>
    [XmlAttribute("num")] List,

    /// <summary>
    /// Specifies that the current tab stop shall result in a location in the document where
    /// all following text is right aligned (i.e. all text runs following this tab stop and
    /// preceding the next tab stop shall be right aligned with respect to the tab stop location).
    /// </summary>
    [XmlAttribute("right")] Right
}
