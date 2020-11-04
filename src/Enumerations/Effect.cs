using System.Xml.Serialization;

namespace DXPlus
{
    public enum Effect
    {
        None,
        Shadow,
        Outline,
        OutlineShadow,
        Emboss,
        [XmlAttribute("imprint")]
        Engrave
    };
}