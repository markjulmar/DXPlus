using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Underline style for text
/// </summary>
public enum UnderlineStyle
{
    /// <summary>
    /// No underline
    /// </summary>
    None = 0,

    /// <summary>
    /// Single line
    /// </summary>
    [XmlAttribute("single")] SingleLine = 1,

    /// <summary>
    /// Underline words separately
    /// </summary>
    Words = 2,

    /// <summary>
    /// Double-line
    /// </summary>
    [XmlAttribute("double")] DoubleLine = 3,

    /// <summary>
    /// Dotted line
    /// </summary>
    Dotted = 4,

    /// <summary>
    /// Thick line
    /// </summary>
    Thick = 6,

    /// <summary>
    /// Dashed line
    /// </summary>
    Dash = 7,

    /// <summary>
    /// Dotted-dash line
    /// </summary>
    DotDash = 9,

    /// <summary>
    /// Dot-dot dash line
    /// </summary>
    DotDotDash = 10,

    /// <summary>
    /// Wavy line
    /// </summary>
    Wave = 11,

    /// <summary>
    /// Heavy dotted line
    /// </summary>
    DottedHeavy = 20,

    /// <summary>
    /// Heavy dashed line
    /// </summary>
    DashedHeavy = 23,

    /// <summary>
    /// Heavy dash/dot line
    /// </summary>
    DashDotHeavy = 25,

    /// <summary>
    /// Heavy dash/dot/dot line
    /// </summary>
    DashDotDotHeavy = 26,

    /// <summary>
    /// Heavy long dash
    /// </summary>
    DashLongHeavy = 27,

    /// <summary>
    /// Long dash
    /// </summary>
    DashLong = 39,

    /// <summary>
    /// Double-wavy line
    /// </summary>
    WavyDouble = 43,

    /// <summary>
    /// Heavy wavy line
    /// </summary>
    WavyHeavy = 55,
};