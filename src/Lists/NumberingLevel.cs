using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a single level (w:lvl) in a numbering definition style.
/// </summary>
public sealed class NumberingLevel : XElementWrapper, IEquatable<NumberingLevel>
{
    private const int Indent = 720;

    /// <summary>
    /// The XML fragment making up this object (w:lvl)
    /// </summary>
    private new XElement Xml => base.Xml!;

    /// <summary>
    /// Level this definition is associated with
    /// </summary>
    public int Level
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "ilvl"), out var result) ? result : -1;
        internal set => Xml.SetAttributeValue(Namespace.Main + "ilvl", value);
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
    /// Specifies a one-based index which determines when a numbering level
    /// should restart to its Start value. A numbering level restarts when an
    /// instance of the specified numbering level is used in the given document's contents.
    /// </summary>
    public int RestartAt
    {
        get => int.Parse(Xml.Element(Namespace.Main + "lvlRestart")?.GetVal() ?? "0");
        set => Xml.AddElementVal(Namespace.Main + "lvlRestart", value<1 ? 1 : value);
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
    /// Public constructor to add new levels/overrides.
    /// </summary>
    public NumberingLevel(int level, NumberingFormat format) : this(new XElement(Namespace.Main + "lvl"))
    {
        Level = level;
        Start = 1;
        Format = format;
        Text = GetDefaultForFormat(format, level);
        Alignment = Alignment.Left;
        ParagraphFormatting.LeftIndent = Indent * level + Indent;
        ParagraphFormatting.HangingIndent = Indent / 2;
    }

    /// <summary>
    /// Returns the default text for the specified format.
    /// </summary>
    /// <param name="format">Format</param>
    /// <param name="level">Level</param>
    /// <returns>Default text</returns>
    private static string GetDefaultForFormat(NumberingFormat format, int level)
    {
        switch (format)
        {
            case NumberingFormat.Bullet: 
                return level == 0 ? "•" : "-";
            case NumberingFormat.DecimalEnclosedCircle:
            case NumberingFormat.DecimalEnclosedFullstop:
            case NumberingFormat.DecimalEnclosedParen:
            case NumberingFormat.DecimalEnclosedCircleChinese:
                return $"(%{level})";
            case NumberingFormat.UpperRoman:
            case NumberingFormat.UpperLetter:
            case NumberingFormat.LowerLetter:
            case NumberingFormat.LowerRoman:
            case NumberingFormat.RussianLower:
            case NumberingFormat.RussianUpper:
                return $"%{level})";
            case NumberingFormat.Ordinal:
            case NumberingFormat.CardinalText:
            case NumberingFormat.OrdinalText:
            case NumberingFormat.Hex:
            case NumberingFormat.Chicago:
            case NumberingFormat.Numbered:
            case NumberingFormat.DecimalFullWidth:
            case NumberingFormat.DecimalHalfWidth:
            case NumberingFormat.DecimalFullWidth2:
            case NumberingFormat.DecimalZero:
                return $"%{level}.";
            case NumberingFormat.NumberInDash:
                return $"- %{level}";
            case NumberingFormat.Removed:
                throw new ArgumentException("Removed is not a valid numbering format.");
            default:
                return "%{level}";
        }
    }

    /// <summary>
    /// Public constructor to add new levels/overrides.
    /// </summary>
    internal NumberingLevel(int level) : this(new XElement(Namespace.Main + "lvl"))
    {
        Level = level;
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="xml"></param>
    internal NumberingLevel(XElement xml)
    {
        base.Xml = xml;
    }

    /// <summary>Indicates whether the current object is equal to another object of the same type.</summary>
    /// <param name="other">An object to compare with this object.</param>
    /// <returns>
    /// <see langword="true" /> if the current object is equal to the <paramref name="other" /> parameter; otherwise, <see langword="false" />.</returns>
    public bool Equals(NumberingLevel? other)
    {
        return ReferenceEquals(this, other)
               || other != null && XNode.DeepEquals(Xml.Normalize(), other.Xml.Normalize());
    }

    /// <summary>Determines whether the specified object is equal to the current object.</summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>
    /// <see langword="true" /> if the specified object  is equal to the current object; otherwise, <see langword="false" />.</returns>
    public override bool Equals(object? obj) 
        => ReferenceEquals(this, obj) || obj is NumberingLevel other && Equals(other);

    /// <summary>Serves as the default hash function.</summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode() => Xml.GetHashCode();
}