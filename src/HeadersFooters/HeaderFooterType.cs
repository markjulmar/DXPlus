using System.Xml.Serialization;

namespace DXPlus
{
    /// <summary>
    /// Header and footer placement
    /// </summary>
    public enum HeaderFooterType
    {
        /// <summary>
        /// Header/Footer on first page.
        /// </summary>
        First,
        /// <summary>
        /// Header/Footer on even pages.
        /// </summary>
        Even,
        /// <summary>
        /// Header/Footer on odd pages.
        /// </summary>
        [XmlAttribute("default")] Odd,
    }
}