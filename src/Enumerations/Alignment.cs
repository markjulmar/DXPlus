using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Text alignment of a Paragraph or List
/// </summary>
public enum Alignment
{
    /// <summary>
    /// Left
    /// </summary>
    Left,
    /// <summary>
    /// Center
    /// </summary>
    Center,
    /// <summary>
    /// Right
    /// </summary>
    Right,
    /// <summary>
    /// Both
    /// </summary>
    Both,
    /// <summary>
    /// Distribute evenly
    /// </summary>
    Distribute,
    /// <summary>
    /// Aligned to list
    /// </summary>
    [XmlAttribute("numTab")] AlignToListTab,
    /// <summary>
    /// Medium Kashida
    /// </summary>
    MediumKashida,
    /// <summary>
    /// High Kashida
    /// </summary>
    HighKashida,
    /// <summary>
    /// Low Kashida
    /// </summary>
    LowKashida,
    /// <summary>
    /// Thai Distributed
    /// </summary>
    ThaiDistribute
};