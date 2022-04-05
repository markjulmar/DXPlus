using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Table Cell and Paragraph Border styles
/// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
/// </summary>
public enum BorderStyle
{
    /// <summary>
    /// No border
    /// </summary>
    [XmlAttribute("nil")] None = 0,
    /// <summary>
    /// Single line border
    /// </summary>
    Single,
    /// <summary>
    /// Thick line border
    /// </summary>
    Thick,
    /// <summary>
    /// Double-line border
    /// </summary>
    Double,
    /// <summary>
    /// Dotted line border
    /// </summary>
    Dotted,
    /// <summary>
    /// Dashed line border
    /// </summary>
    Dashed,
    /// <summary>
    /// Dot-dash line border
    /// </summary>
    DotDash,
    /// <summary>
    /// Dot-dot-dash line border
    /// </summary>
    DotDotDash,
    /// <summary>
    /// Triple line border
    /// </summary>
    Triple,
    /// <summary>
    /// Thin-thick with small gap
    /// </summary>
    ThinThickSmallGap,
    /// <summary>
    /// Thick-thin with small gap
    /// </summary>
    ThickThinSmallGap,
    /// <summary>
    /// Thin-thick-thin with small gap
    /// </summary>
    ThinThickThinSmallGap,
    /// <summary>
    /// Thin-thick with medium gap
    /// </summary>
    ThinThickMediumGap,
    /// <summary>
    /// Thick-thin with medium gap
    /// </summary>
    ThickThinMediumGap,
    /// <summary>
    /// Thin-thick-thin with medium gap
    /// </summary>
    ThinThickThinMediumGap,
    /// <summary>
    /// Thin-thick with large gap
    /// </summary>
    ThinThickLargeGap,
    /// <summary>
    /// Thick-thin with large gap
    /// </summary>
    ThickThinLargeGap,
    /// <summary>
    /// Thin-thick-thick with large gap
    /// </summary>
    ThinThickThinLargeGap,
    /// <summary>
    /// Wavy line
    /// </summary>
    Wave,
    /// <summary>
    /// Double-wavy line
    /// </summary>
    DoubleWave,
    /// <summary>
    /// Dash with small gap
    /// </summary>
    DashSmallGap,
    /// <summary>
    /// Dash-dot stroked line
    /// </summary>
    DashDotStroked,
    /// <summary>
    /// 3D embossed
    /// </summary>
    ThreeDEmboss,
    /// <summary>
    /// 3D engraved
    /// </summary>
    ThreeDEngrave,
    /// <summary>
    /// Outset
    /// </summary>
    Outset,
    /// <summary>
    /// Inset
    /// </summary>
    Inset
}