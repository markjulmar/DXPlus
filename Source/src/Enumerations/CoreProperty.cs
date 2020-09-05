using System.Xml.Serialization;

namespace DXPlus
{
    /// <summary>
    /// Available document properties stored in core.xml
    /// </summary>
    public enum DocumentPropertyName
    {
        [XmlAttribute("dc:title")]
        Title,
        [XmlAttribute("dc:subject")]
        Subject,
        [XmlAttribute("dc:creator")]
        Creator,
        [XmlAttribute("cp:keywords")]
        Keywords,
        [XmlAttribute("dc:description")]
        Comments,
        [XmlAttribute("cp:lastModifiedBy")]
        LastSavedBy,
        [XmlAttribute("cp:revision")]
        Revision,
        [XmlAttribute("cp:category")]
        Category,
        [XmlAttribute("dcterms:created")]
        CreatedDate,
        [XmlAttribute("dcterms:modified")]
        SaveDate
    }
}
