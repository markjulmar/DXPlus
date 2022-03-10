using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Text effects
/// </summary>
public enum Effect
{
    /// <summary>
    /// None
    /// </summary>
    None,

    /// <summary>
    /// Shadow
    /// </summary>
    Shadow,

    /// <summary>
    /// Outline
    /// </summary>
    Outline,

    /// <summary>
    /// Outline + shadow
    /// </summary>
    OutlineShadow,

    /// <summary>
    /// Embossed
    /// </summary>
    Emboss,

    /// <summary>
    /// Engraved
    /// </summary>
    [XmlAttribute("imprint")] Engrave
};