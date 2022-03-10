using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Represents the switches set on a TOC.
/// See https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_TOCTOC_topic_ID0ELZO1.html
/// </summary>
[Flags]
public enum TableOfContentsSwitches
{
    /// <summary>
    /// None
    /// </summary>
    None = 0,

    /// <summary>
    /// Includes captioned items, but omits caption labels and numbers.
    /// </summary>
    [XmlAttribute(@"\a")] A = 1,

    /// <summary>
    /// Includes entries only from the named bookmark.
    /// </summary>
    [XmlAttribute(@"\b")] B = 2,
    
    /// <summary>
    /// Include figures, charts and tables.
    /// </summary>
    [XmlAttribute(@"\c")] C = 4,
    
    /// <summary>
    /// Separator between sequence and page#. Default is hyphen
    /// </summary>
    [XmlAttribute(@"\d")] D = 8,
    
    /// <summary>
    /// Include only the matching identifiers
    /// </summary>
    [XmlAttribute(@"\f")] F = 16,
    
    /// <summary>
    /// Uses hyperlinks
    /// </summary>
    [XmlAttribute(@"\h")] H = 32,
    
    /// <summary>
    /// Include specific levels
    /// </summary>
    [XmlAttribute(@"\l")] L = 64,
    
    /// <summary>
    /// Omit page numbers
    /// </summary>
    [XmlAttribute(@"\n")] N = 128,
    
    /// <summary>
    /// Use paragraphs formatted with heading styles
    /// </summary>
    [XmlAttribute(@"\o")] O = 256,
    
    /// <summary>
    /// Separator between entry and page#. Default is tab
    /// </summary>
    [XmlAttribute(@"\p")] P = 512,
    
    /// <summary>
    /// Sequences get an added prefix to the page#.
    /// </summary>
    [XmlAttribute(@"\s")] S = 1024,
    
    /// <summary>
    /// Use the specific paragraph styles
    /// </summary>
    [XmlAttribute(@"\t")] T = 2048,
    
    /// <summary>
    /// Use the applied paragraph outline level
    /// </summary>
    [XmlAttribute(@"\u")] U = 4096,
    
    /// <summary>
    /// Preserve tab entries within table entries
    /// </summary>
    [XmlAttribute(@"\w")] W = 8192,
    
    /// <summary>
    /// Preserve newline characters within table entries
    /// </summary>
    [XmlAttribute(@"\x")] X = 16384,
    
    /// <summary>
    /// Hides tab leader and page numbers in Web layout view
    /// </summary>
    [XmlAttribute(@"\z")] Z = 32768
}