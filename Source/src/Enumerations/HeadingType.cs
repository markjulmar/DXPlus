using System.Xml.Serialization;

namespace DXPlus
{
    public enum HeadingType
    {
        [XmlAttribute("Heading1")]
        Heading1,

        [XmlAttribute("Heading2")]
        Heading2,

        [XmlAttribute("Heading3")]
        Heading3,

        [XmlAttribute("Heading4")]
        Heading4,

        [XmlAttribute("Heading5")]
        Heading5,

        [XmlAttribute("Heading6")]
        Heading6,

        [XmlAttribute("Heading7")]
        Heading7,

        [XmlAttribute("Heading8")]
        Heading8,

        [XmlAttribute("Heading9")]
        Heading9,

        [XmlAttribute("NoSpacing")]
        NoSpacing,

        [XmlAttribute("Title")]
        Title,

        [XmlAttribute("Subtitle")]
        Subtitle,

        [XmlAttribute("Quote")]
        Quote,

        [XmlAttribute("IntenseQuote")]
        IntenseQuote,

        [XmlAttribute("Emphasis")]
        Emphasis,

        [XmlAttribute("IntenseEmphasis")]
        IntenseEmphasis,

        [XmlAttribute("Strong")]
        Strong,

        [XmlAttribute("ListParagraph")]
        ListParagraph,

        [XmlAttribute("SubtleReference")]
        SubtleReference,

        [XmlAttribute("IntenseReference")]
        IntenseReference,
    }
}