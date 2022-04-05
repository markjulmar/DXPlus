using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Specifies the type of numbering allowed in a definition. Note that this
/// restriction is applied only to the UI - it's possible to have multi-level lists
/// even when the level type is defined as 'single'.
/// </summary>
public enum NumberingLevelType
{
    /// <summary>
    /// Numbering format has only one level
    /// </summary>
    [XmlAttribute("singleLevel")] Single,
        
    /// <summary>
    /// Multiple levels of the same type (numbers, bullets, etc.)
    /// </summary>
    [XmlAttribute("multiLevel")] Multi,
        
    /// <summary>
    /// Hybrid list of multiple levels of different types.
    /// </summary>
    [XmlAttribute("hybridMultilevel")] Hybrid
}