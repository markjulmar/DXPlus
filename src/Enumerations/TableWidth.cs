using System.Xml.Serialization;

namespace DXPlus
{
    /// <summary>
    /// Specifies the units of the table width property.
    /// </summary>
    public enum TableWidthUnit
    {
        /// <summary>
        /// Specifies that width is determined by the overall table layout algorithm.
        /// </summary>
        Auto,
        /// <summary>
        /// Specifies that the value is in twentieths of a point (1/1440 of an inch).
        /// </summary>
        Dxa,
        /// <summary>
        /// Specifies a value of zero.
        /// </summary>
        [XmlAttribute("nil")]
        None,
        /// <summary>
        /// Specifies a value as a percent of the table width.
        /// </summary>
        [XmlAttribute("pct")]
        Percentage
    }
}
