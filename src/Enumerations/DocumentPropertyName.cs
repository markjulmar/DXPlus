using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Available document properties stored in core.xml
/// </summary>
public enum DocumentPropertyName
{
    /// <summary>
    /// Title
    /// </summary>
    [XmlAttribute("dc:title")] Title,
        
    /// <summary>
    /// Subject
    /// </summary>
    [XmlAttribute("dc:subject")] Subject,

    /// <summary>
    /// The creator
    /// </summary>
    [XmlAttribute("dc:creator")] Creator,

    /// <summary>
    /// Keywords
    /// </summary>
    [XmlAttribute("cp:keywords")] Keywords,

    /// <summary>
    /// Description/Comments
    /// </summary>
    [XmlAttribute("dc:description")] Description,

    /// <summary>
    /// Last modified by author
    /// </summary>
    [XmlAttribute("cp:lastModifiedBy")] LastSavedBy,

    /// <summary>
    /// Revision/Version
    /// </summary>
    [XmlAttribute("cp:revision")] Revision,

    /// <summary>
    /// Category
    /// </summary>
    [XmlAttribute("cp:category")] Category,

    /// <summary>
    /// Name of author who created document
    /// </summary>
    [XmlAttribute("dcterms:created")] CreatedDate,

    /// <summary>
    /// Last date/time document was saved
    /// </summary>
    [XmlAttribute("dcterms:modified")] SaveDate,

    /// <summary>
    /// Status of the document (draft, final, etc.)
    /// </summary>
    [XmlAttribute("cp:contentStatus")] Status
}