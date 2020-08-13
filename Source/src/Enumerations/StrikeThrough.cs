using System.Xml.Serialization;

namespace DXPlus
{
    public enum StrikeThrough
    {
        None,
        Strike,
        [XmlAttribute("dstrike")]
        DoubleStrike
    };
}