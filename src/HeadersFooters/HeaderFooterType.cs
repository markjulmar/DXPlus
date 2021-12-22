using System.Xml.Serialization;

namespace DXPlus
{
    public enum HeaderFooterType
    {
        First,
        Even,
        [XmlAttribute("default")]
        Odd,
    }
}