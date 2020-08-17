using System.Xml.Serialization;

namespace DXPlus
{
    public enum Strikethrough
    {
        None,
        Strike,
        [XmlAttribute("dstrike")]
        DoubleStrike
    };
}