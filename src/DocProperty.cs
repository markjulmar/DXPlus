using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a field in the document.
/// This field displays the value stored in a document or custom property.
/// </summary>
public sealed class DocProperty : DocXElement, IEquatable<DocProperty>
{
    private const string DocPropertyText = "DOCPROPERTY";

    /// <summary>
    /// Name of the property
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Value of the property
    /// </summary>
    public string? Value { get; }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="document"></param>
    /// <param name="packagePart"></param>
    /// <param name="name"></param>
    /// <param name="value"></param>
    internal DocProperty(Document document, PackagePart packagePart, XElement name, XElement? value) : base(name)
    {
        SetOwner(document, packagePart, false);

        var dpre = new Regex($"{DocPropertyText} (?<name>.*) \\*");

        // Check for a simple field
        string? instr = Xml.AttributeValue(Internal.Name.Instr, null)?.Trim();
        if (instr != null)
        {
            Name = instr.Contains(DocPropertyText)
                ? dpre.Match(instr).Groups["name"].Value.Trim('"')
                : instr[..instr.IndexOf(' ')];
            Value = Xml.Descendants().First(e => e.Name == Internal.Name.Text)?.Value;
        }
        // Complex field
        else
        {
            if (value == null) throw new ArgumentNullException(nameof(value));

            var instrText = Xml.Descendants(Namespace.Main + "instrText").Single();
            string text = instrText.Value.Trim();
            Name = text.Contains(DocPropertyText)
                ? dpre.Match(text).Groups["name"].Value.Trim('"')
                : text[..text.IndexOf(' ')];
            Value = new Run(document, packagePart, value, 0).Text;
        }
    }

    /// <summary>
    /// Determines equality for doc properties
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(DocProperty? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for this property
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as DocProperty);

    /// <summary>
    /// Returns hashcode for this property
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();
}