using System.Xml.Serialization;

namespace DXPlus
{
    /// <summary>
    /// Text alignment of a Paragraph or List
    /// </summary>
    public enum Alignment
    {
        Left,
        Center,
        Right,
        Both,
        Distribute,
        [XmlAttribute("numTab")]
        AlignToListTab,
        MediumKashida,
        HighKashida,
        LowKashida,
        ThaiDistribute
    };
}