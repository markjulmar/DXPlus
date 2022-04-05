using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Specifies a set of properties which shall dictate the appearance and behavior 
/// of a set of numbered paragraphs in a Word document. This is persisted as a w:abstractNum object
/// in the /word/numbering.xml document.
/// </summary>
public sealed class NumberingStyle : XElementWrapper, IEquatable<NumberingStyle>
{
    private new XElement Xml => base.Xml!;
    private readonly XElementCollection<NumberingLevel> levels;

    /// <summary>
    /// Specifies a unique number which will be used as the identifier for the numbering definition.
    /// The value is referenced by numbering instances (num) via the num's abstractNumId child element.
    /// </summary>
    public int Id
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "abstractNumId"), out var result) ? result : -1;
        internal set => Xml.SetAttributeValue(Namespace.Main + "abstractNumId", value);
    }

    /// <summary>
    /// Unique hexadecimal identifier for the numbering style. This value will be
    /// the same for two numbering styles based on the same initial definition (e.g.
    /// where a new definition is created from an existing one). This is persisted as w:nsid.
    /// </summary>
    public uint? CreatorId
    {
        get => uint.TryParse(Xml.Element(Namespace.Main + "nsid").GetVal(null), out var result) ? result : null;
        set
        {
            if (value == null)
            {
                Xml.Element(Namespace.Main + "nsid")?.Remove();
            }
            else
            {
                if (!DocumentHelpers.IsValidHexNumber(value.ToString()??""))
                    throw new ArgumentException("Invalid hex value.", nameof(CreatorId));
                Xml.AddElementVal(Namespace.Main + "nsid", value);
            }
        }
    }

    /// <summary>
    /// The type of numbering shown in the UI for lists using this definition - single, multi-level, etc.
    /// </summary>
    public NumberingLevelType? LevelType
    {
        get => Xml.Element(Namespace.Main + "multiLevelType").GetVal()
            .TryGetEnumValue<NumberingLevelType>(out var result) ? result : null;
        set => Xml.AddElementVal(Namespace.Main + "multiLevelType", value?.GetEnumName());
    }

    /// <summary>
    /// Optional user-friendly name (alias) for this numbering definition.
    /// </summary>
    public string? Name
    {
        get => Xml.Element(Internal.Name.NameId).GetVal(null);
        set => Xml.AddElementVal(Internal.Name.NameId, string.IsNullOrWhiteSpace(value) ? null : value);
    }

    /// <summary>
    /// Optional style name that indicates this definition doesn't contain any style properties but
    /// uses the specified style for all properties.
    /// </summary>
    public string? NumStyleLink
    {
        get => Xml.Element(Namespace.Main + "numStyleLink").GetVal(null);
        set => Xml.AddElementVal(Namespace.Main + "numStyleLink", string.IsNullOrWhiteSpace(value) ? null : value);
    }

    /// <summary>
    /// Specifies this definition is the base numbering definition for the reference numbering style.
    /// </summary>
    public string? StyleLink
    {
        get => Xml.Element(Namespace.Main + "styleLink").GetVal(null);
        set => Xml.AddElementVal(Namespace.Main + "styleLink", string.IsNullOrWhiteSpace(value) ? null : value);
    }

    /// <summary>
    /// The levels
    /// </summary>
    public IList<NumberingLevel> Levels => levels;

    /// <summary>
    /// Public constructor
    /// </summary>
    public NumberingStyle() : this(new XElement(Namespace.Main + "abstractNum"))
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="xml">XML definition (abstractNum)</param>
    internal NumberingStyle(XElement xml)
    {
        base.Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        levels = new XElementCollection<NumberingLevel>(Xml, null, Namespace.Main + "lvl",
            element => new NumberingLevel(element));
    }

    /// <summary>Indicates whether the current object is equal to another object of the same type.</summary>
    /// <param name="other">An object to compare with this object.</param>
    /// <returns>
    /// <see langword="true" /> if the current object is equal to the <paramref name="other" /> parameter; otherwise, <see langword="false" />.</returns>
    public bool Equals(NumberingStyle? other)
    {
        return ReferenceEquals(this, other)
               || other != null && XNode.DeepEquals(Xml.Normalize(), other.Xml.Normalize());
    }

    /// <summary>Determines whether the specified object is equal to the current object.</summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>
    /// <see langword="true" /> if the specified object  is equal to the current object; otherwise, <see langword="false" />.</returns>
    public override bool Equals(object? obj)
        => ReferenceEquals(this, obj) || obj is NumberingStyle other && Equals(other);

    /// <summary>Serves as the default hash function.</summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode() => Xml.GetHashCode();
}