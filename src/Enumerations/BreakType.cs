using System.Xml.Serialization;

namespace DXPlus
{
    /// <summary>
    /// Break types that can be inserted into a text run.
    /// </summary>
    public enum BreakType
    {
        /// <summary>
        /// Specifies that the current break shall restart itself on the next page
        /// of the document when the document is displayed in page view.
        /// </summary>
        Page,
        /// <summary>
        /// Specifies that the current break shall restart itself on the next column
        /// available on the current page when the document is displayed in page view.
        /// </summary>
        Column,
        /// <summary>
        /// Specifies that the current break shall restart itself on the next line
        /// in the document when the document is displayed in page view.
        /// </summary>
        [XmlAttribute("textWrapping")]
        Line
    }
}
