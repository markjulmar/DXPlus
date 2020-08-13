using System.Xml.Serialization;

namespace DXPlus
{

    public enum UnderlineStyle
    {
        None = 0,
        [XmlAttribute("single")]
        SingleLine = 1,
        Words = 2,
        [XmlAttribute("double")]
        DoubleLine = 3,
        Dotted = 4,
        Thick = 6,
        Dash = 7,
        DotDash = 9,
        DotDotDash = 10,
        Wave = 11,
        DottedHeavy = 20,
        DashedHeavy = 23,
        DashDotHeavy = 25,
        DashDotDotHeavy = 26,
        DashLongHeavy = 27,
        DashLong = 39,
        WavyDouble = 43,
        WavyHeavy = 55,
    };
}