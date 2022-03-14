using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Built-in styles
/// </summary>
public enum HeadingType
{
    /// <summary>
    /// Heading1
    /// </summary>
    [XmlAttribute("Heading1")] Heading1,

    /// <summary>
    /// Heading2
    /// </summary>
    [XmlAttribute("Heading2")] Heading2,

    /// <summary>
    /// Heading3
    /// </summary>
    [XmlAttribute("Heading3")] Heading3,

    /// <summary>
    /// Heading4
    /// </summary>
    [XmlAttribute("Heading4")] Heading4,

    /// <summary>
    /// Heading5
    /// </summary>
    [XmlAttribute("Heading5")] Heading5,

    /// <summary>
    /// Heading6
    /// </summary>
    [XmlAttribute("Heading6")] Heading6,

    /// <summary>
    /// Heading7
    /// </summary>
    [XmlAttribute("Heading7")] Heading7,

    /// <summary>
    /// Heading8
    /// </summary>
    [XmlAttribute("Heading8")] Heading8,

    /// <summary>
    /// Heading9
    /// </summary>
    [XmlAttribute("Heading9")] Heading9,

    /// <summary>
    /// NoSpacing
    /// </summary>
    [XmlAttribute("NoSpacing")] NoSpacing,

    /// <summary>
    /// Title
    /// </summary>
    [XmlAttribute("Title")] Title,

    /// <summary>
    /// Subtitle
    /// </summary>
    [XmlAttribute("Subtitle")] Subtitle,

    /// <summary>
    /// Quote
    /// </summary>
    [XmlAttribute("Quote")] Quote,

    /// <summary>
    /// IntenseQuote
    /// </summary>
    [XmlAttribute("IntenseQuote")] IntenseQuote,

    /// <summary>
    /// Emphasis
    /// </summary>
    [XmlAttribute("Emphasis")] Emphasis,

    /// <summary>
    /// IntenseEmphasis
    /// </summary>
    [XmlAttribute("IntenseEmphasis")] IntenseEmphasis,

    /// <summary>
    /// Strong
    /// </summary>
    [XmlAttribute("Strong")] Strong,

    /// <summary>
    /// ListParagraph
    /// </summary>
    [XmlAttribute("ListParagraph")] ListParagraph,

    /// <summary>
    /// SubtleReference
    /// </summary>
    [XmlAttribute("SubtleReference")] SubtleReference,

    /// <summary>
    /// IntenseReference
    /// </summary>
    [XmlAttribute("IntenseReference")] IntenseReference,

    /// <summary>
    /// Caption
    /// </summary>
    [XmlAttribute("Caption")] Caption,
}