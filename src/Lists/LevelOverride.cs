using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents an override section in a numbering definition.
/// </summary>
public sealed class LevelOverride
{
    /// <summary>
    /// The XML fragment making up this object (w:lvlOverride)
    /// </summary>
    private XElement Xml { get; }

    /// <summary>
    /// Level to override
    /// </summary>
    public int Level
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "ilvl"), out var result) ? result : throw new DocumentFormatException(nameof(Level));
        set => Xml.SetAttributeValue(Namespace.Main + "ilvl", value);
    }

    /// <summary>
    /// Specifies the number which the specified level override shall begin with.
    /// This value is used when this level initially starts in a document, as well as whenever it is restarted
    /// </summary>
    public int? Start
    {
        get => int.TryParse(Xml.Element(Namespace.Main + "startOverride").GetVal(), out var result) ? result : null;
        set => Xml.AddElementVal(Namespace.Main + "startOverride", value?.ToString());
    }

    /// <summary>
    /// Returns the details of the overriden level
    /// </summary>
    public NumberingLevel NumberingLevel => new(Xml.Element(Namespace.Main + "lvl") ?? throw new DocumentFormatException(nameof(NumberingLevel)));

    /// <summary>
    /// Removes this level override
    /// </summary>
    public void Remove()
    {
        Xml.Remove();
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="xml"></param>
    internal LevelOverride(XElement xml)
    {
        Xml = xml ?? throw new ArgumentNullException(nameof(xml));
    }
}