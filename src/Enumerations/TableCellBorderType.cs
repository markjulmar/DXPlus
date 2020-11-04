using System.Xml.Serialization;

namespace DXPlus
{

    /// <summary>
    /// Table Cell Border Types
    /// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
    /// </summary>
    public enum TableCellBorderType
    {
        Top,
        Bottom,
        Left,
        Right,
        InsideH,
        InsideV,
        [XmlAttribute("tl2br")]
        TopLeftToBottomRight,
        [XmlAttribute("tr2bl")]
        TopRightToBottomLeft
    }
}