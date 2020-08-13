using System.Xml.Serialization;

namespace DXPlus
{
    public enum Misc
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