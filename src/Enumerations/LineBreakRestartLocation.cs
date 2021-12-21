using System.Xml.Serialization;

namespace DXPlus
{
    public enum LineBreakRestartLocation
    {
        /// <summary>
        /// Specifies that the text wrapping break shall advance the text to the next line
        /// in the document which spans the full width of the line
        /// (i.e. the next line which is not interrupted by any floating objects
        /// when those objects are positioned on the page at display time.
        /// </summary>
        All,
        /// <summary>
        /// Restart In next text region When In Leftmost Position.
        /// </summary>
        Left,
        /// <summary>
        /// Restart In next text region When In Rightmost Position.
        /// </summary>
        Right,
        /// <summary>
        /// Specifies that the text wrapping break shall advance the text to the next line
        /// in the document, regardless of its position left to right or the presence
        /// of any floating objects which intersect with the line. This is the default
        /// </summary>
        [XmlAttribute("none")]
        NextLine
    }
}
