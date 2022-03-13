using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Identified sections in a document
/// </summary>
public sealed class Section : DocXElement, IEquatable<Section>
{
    /// <summary>
    /// Returns a collection of Headers in this section of the document.
    /// A document typically contains three Headers.
    /// A default one (odd), one for the first page and one for even pages.
    /// </summary>
    public HeaderCollection Headers { get; }

    /// <summary>
    /// Returns a collection of Footers in this section of the document.
    /// A document typically contains three Footers.
    /// A default one (odd), one for the first page and one for even pages.
    /// </summary>
    public FooterCollection Footers { get; }

    /// <summary>
    /// Paragraphs contained in this section.
    /// </summary>
    public IEnumerable<Paragraph> Paragraphs
    {
        get
        {
            // Get all paragraphs from the owner document.
            var paragraphs = Document.Paragraphs.ToList();

            Paragraph? startingParagraph;

            // If this is the final section
            if (Xml.Name == Namespace.Main + "body")
            {
                startingParagraph = paragraphs.LastOrDefault();
            }
            else
            {
                // Locate this section in the paragraphs.
                startingParagraph = paragraphs
                    .Where(p => p.Xml.Element(Name.ParagraphProperties, Name.SectionProperties) != null)
                    .SingleOrDefault(p => XNode.DeepEquals(p.Xml.Normalize(), Xml.Normalize()));
            }

            if (startingParagraph == null)
            {
                return Enumerable.Empty<Paragraph>();
            }

            var sectionParagraphs = new List<Paragraph> {startingParagraph};

            // Starting at the current element, walk backward until we hit a paragraph
            // with section properties or the parent.
            int index = paragraphs.IndexOf(startingParagraph);
            for (int i = index-1; i >= 0; i--)
            {
                var p = paragraphs[i];
                if (p.Xml.Element(Name.ParagraphProperties, Name.SectionProperties) != null)
                    break;
                sectionParagraphs.Add(p);
            }

            // Reverse it.
            sectionParagraphs.Reverse();
            return sectionParagraphs;
        }
    }

    /// <summary>
    /// Properties tied to this section
    /// </summary>
    public SectionProperties Properties { get; }

    /// <summary>
    /// Create a new section
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart"></param>
    /// <param name="xml">Parent of sectionProps</param>
    internal Section(Document document, PackagePart packagePart, XElement xml) : base(xml)
    {
        var eProps = xml.Element(Name.SectionProperties) ??
                     xml.Element(Name.ParagraphProperties, Name.SectionProperties);
        if (eProps == null) throw new ArgumentNullException(nameof(xml));
        Properties = new SectionProperties(eProps);

        SetOwner(document, packagePart, false);

        // Load headers/footers
        Headers = new HeaderCollection(document, this);
        Footers = new FooterCollection(document, this);
    }

    /// <summary>
    /// Determines equality for a section
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Section? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a section
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Section);

    /// <summary>
    /// Returns hashcode for this section
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();
}