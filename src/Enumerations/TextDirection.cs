using System.Xml.Serialization;

namespace DXPlus
{
    public enum TextDirection
    {
        [XmlAttribute("btLr")]
        BottomToTopLeftToEnd,
        [XmlAttribute("lrTb")]
        LeftToRightTopToBottom,
        [XmlAttribute("lrTbV")]
        LeftToRightTopToBottomVertical,
        [XmlAttribute("tbLrV")]
        TopToBottomLeftToRightVertical,
        [XmlAttribute("tbRl")]
        TopToBottomRightToLeft,
        [XmlAttribute("tbRlV")]
        TopToBottomRightToLeftVertical,
    };
}