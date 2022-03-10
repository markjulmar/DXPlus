using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Shading pattern applied to a paragraph, table cell, or border.
/// </summary>
public enum ShadePattern
{
    /// <summary>
    /// Transparent
    /// </summary>
    Clear,

    /// <summary>
    /// Diagonal cross
    /// </summary>
    [XmlAttribute("diagCross")] DiagonalCross,

    /// <summary>
    /// Diagonal stripe
    /// </summary>
    [XmlAttribute("diagStripe")] DiagonalStripe,

    /// <summary>
    /// Horizontal cross
    /// </summary>
    [XmlAttribute("horzCross")] HorizontalCross,

    /// <summary>
    /// Horizontal stripe
    /// </summary>
    [XmlAttribute("horzStripe")] HorizontalStripe,

    /// <summary>
    /// No shading
    /// </summary>
    [XmlAttribute("nil")] None,

    /// <summary>
    /// 5%
    /// </summary>
    [XmlAttribute("pct5")] Percent5,

    /// <summary>
    /// 10%
    /// </summary>
    [XmlAttribute("pct10")] Percent10,

    /// <summary>
    /// 12%
    /// </summary>
    [XmlAttribute("pct12")] Percent12,

    /// <summary>
    /// 15%
    /// </summary>
    [XmlAttribute("pct15")] Percent15,

    /// <summary>
    /// 20%
    /// </summary>
    [XmlAttribute("pct20")] Percent20,

    /// <summary>
    /// 25%
    /// </summary>
    [XmlAttribute("pct25")] Percent25,

    /// <summary>
    /// 30%
    /// </summary>
    [XmlAttribute("pct30")] Percent30,

    /// <summary>
    /// 35%
    /// </summary>
    [XmlAttribute("pct35")] Percent35,

    /// <summary>
    /// 37%
    /// </summary>
    [XmlAttribute("pct37")] Percent37,

    /// <summary>
    /// 40%
    /// </summary>
    [XmlAttribute("pct40")] Percent40,

    /// <summary>
    /// 45%
    /// </summary>
    [XmlAttribute("pct45")] Percent45,
    
    /// <summary>
    /// 50%
    /// </summary>
    [XmlAttribute("pct50")] Percent50,
    
    /// <summary>
    /// 55%
    /// </summary>
    [XmlAttribute("pct55")] Percent55,

    /// <summary>
    /// 60%
    /// </summary>
    [XmlAttribute("pct60")] Percent60,

    /// <summary>
    /// 62%
    /// </summary>
    [XmlAttribute("pct62")] Percent62,
    
    /// <summary>
    /// 65%
    /// </summary>
    [XmlAttribute("pct65")] Percent65,
    
    /// <summary>
    /// 70%
    /// </summary>
    [XmlAttribute("pct70")] Percent70,
    
    /// <summary>
    /// 75%
    /// </summary>
    [XmlAttribute("pct75")] Percent75,

    /// <summary>
    /// 80%
    /// </summary>
    [XmlAttribute("pct80")] Percent80,

    /// <summary>
    /// 85%
    /// </summary>
    [XmlAttribute("pct85")] Percent85,

    /// <summary>
    /// 87%
    /// </summary>
    [XmlAttribute("pct87")] Percent87,

    /// <summary>
    /// 90%
    /// </summary>
    [XmlAttribute("pct90")] Percent90,

    /// <summary>
    /// 95%
    /// </summary>
    [XmlAttribute("pct95")] Percent95,

    /// <summary>
    /// Reverse diagonal stripe
    /// </summary>
    [XmlAttribute("reverseDiagStripe")] ReverseDiagonalStripe,

    /// <summary>
    /// Solid pattern
    /// </summary>
    Solid,

    /// <summary>
    /// Thin diagonal cross
    /// </summary>
    [XmlAttribute("thinDiagCross")] ThinDiagonalCross,

    /// <summary>
    /// Thin diagonal stripe
    /// </summary>
    [XmlAttribute("thinDiagStripe")] ThinDiagonalStripe,

    /// <summary>
    /// Thin horizontal cross
    /// </summary>
    [XmlAttribute("thinHorzCross")] ThinHorizontalCross,

    /// <summary>
    /// Thin horizontal stripe
    /// </summary>
    [XmlAttribute("thinHorzStripe")] ThinHorizontalStripe,

    /// <summary>
    /// Thin reverse diagonal stripe
    /// </summary>
    [XmlAttribute("thinReverseDiagStripe")] ThinReverseDiagonalStripe,

    /// <summary>
    /// Thin vertical stripe
    /// </summary>
    [XmlAttribute("thinVertStripe")] ThinVerticalStripe,

    /// <summary>
    /// Vertical stripe
    /// </summary>
    [XmlAttribute("vertStripe")] VerticalStripe,
}