using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a single level (w:lvl) in a numbering definition style.
/// </summary>
public sealed class NumberingLevel
{
    /// <summary>
    /// The XML fragment making up this object (w:lvl)
    /// </summary>
    internal XElement Xml { get; }

    /// <summary>
    /// Level this definition is associated with
    /// </summary>
    public int Level
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "ilvl"), out var result) ? result : throw new DocumentFormatException(nameof(Level));
        set => Xml.SetAttributeValue(Namespace.Main + "ilvl", value);
    }

    /// <summary>
    /// The starting value for this level.
    /// </summary>
    public int Start
    {
        get => int.Parse(Xml.Element(Namespace.Main + "start")?.GetVal() ?? "0");
        set => Xml.AddElementVal(Namespace.Main + "start", value);
    }

    /// <summary>
    /// specifies a one-based index which determines when a numbering level
    /// should restart to its Start value. A numbering level restarts when an
    /// instance of the specified numbering level is used in the given document's contents.
    /// </summary>
    public int Restart
    {
        get => int.Parse(Xml.Element(Namespace.Main + "lvlRestart")?.GetVal() ?? "0");
        set => Xml.AddElementVal(Namespace.Main + "lvlRestart", value);
    }

    /// <summary>
    /// Retrieve the formatting options
    /// </summary>
    public Formatting Formatting => new(Xml.GetOrAddElement(Name.RunProperties));

    /// <summary>
    /// FirstParagraph properties
    /// </summary>
    public ParagraphProperties ParagraphFormatting => new(Xml.GetOrAddElement(Name.ParagraphProperties));

    /// <summary>
    /// Number format used to display all the values at this level.
    /// </summary>
    public NumberingFormat Format
    {
        get => Xml.Element(Namespace.Main + "numFmt")
            .GetVal().TryGetEnumValue<NumberingFormat>(out var result) ? result : NumberingFormat.None;
        set => Xml.AddElementVal(Namespace.Main + "numFmt", value.GetEnumName());
    }

    /// <summary>
    /// Specifies the content to be displayed when displaying a paragraph at this level.
    /// </summary>
    public string? Text
    {
        get => Xml.Element(Namespace.Main + "lvlText").GetVal();
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Value must be provided.", nameof(Text));
            Xml.GetOrAddElement(Namespace.Main + "lvlText").SetAttributeValue(Name.MainVal, value);
        }
    }

    /// <summary>
    /// Alignment of the text within the list
    /// </summary>
    public Alignment Alignment
    {
        get => Xml.Element(Namespace.Main + "lvlJc")
            .GetVal().TryGetEnumValue<Alignment>(out var result) ? result : Alignment.Left;
        set => Xml.AddElementVal(Namespace.Main + "lvlJc", value.GetEnumName());
    }

    /// <summary>
    /// Removes this level
    /// </summary>
    public void Remove()
    {
        Xml.Remove();
    }

    /// <summary>
    /// Public constructor for externals to add new levels/overrides.
    /// </summary>
    public NumberingLevel()
    {
        this.Xml = new XElement(Namespace.Main + "lvl");
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="xml"></param>
    internal NumberingLevel(XElement xml)
    {
        this.Xml = xml;
    }
}