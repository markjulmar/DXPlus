using System.Xml.Serialization;

namespace DXPlus
{
    public enum ListItemType
    {
        [XmlAttribute("bullet")]
        Bulleted,
        [XmlAttribute("decimal")]
        Numbered,
        [XmlAttribute("decimalEnclosedCircle")]
        CircleNumbered,
        None,
    }
}