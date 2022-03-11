using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Helpers;
using DXPlus.Resources;

namespace DXPlus;

/// <summary>
/// This is a comment tied to one or more text runs.
/// </summary>
public sealed class Comment : DocXElement, IEquatable<Comment>
{
    private const string ParagraphStyle = "CommentText";
    private const string RunStyle = "CommentReference";

    /// <summary>
    /// Unique identifier for the comment, starting at "1"
    /// </summary>
    public int Id
    {
        get => DocumentHelpers.GetId(Xml) ?? throw new DocumentFormatException(nameof(Id));
        init => Xml.SetAttributeValue(Name.Id, value);
    }

    /// <summary>
    /// Author name
    /// </summary>
    public string AuthorName
    {
        get => Xml.AttributeValue(Namespace.Main + "author") ?? throw new DocumentFormatException(nameof(AuthorName), "Missing author on comment.");
        set => Xml.SetAttributeValue(Namespace.Main + "author", value);
    }

    /// <summary>
    /// Author name
    /// </summary>
    public string? AuthorInitials
    {
        get => Xml.AttributeValue(Namespace.Main + "initials");
        set => Xml.SetAttributeValue(Namespace.Main + "initials", value);
    }

    /// <summary>
    /// Author name
    /// </summary>
    public DateTime? Date
    {
        get => DateTime.TryParse(Xml.AttributeValue(Namespace.Main + "date"), out var dt) ? dt : null;
        set => Xml.SetAttributeValue(Namespace.Main + "date", value != null ? value.Value.ToString("s") + "Z" : null);
    }

    /// <summary>
    /// Constructor for the comment
    /// </summary>
    /// <param name="document"></param>
    /// <param name="packagePart"></param>
    /// <param name="authorName">Author</param>
    /// <param name="dateTime">Date</param>
    /// <param name="authorInitials">Initials</param>
    /// <exception cref="ArgumentNullException"></exception>
    internal Comment(Document document, PackagePart packagePart, 
        string authorName, DateTime? dateTime, string? authorInitials)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (packagePart == null) throw new ArgumentNullException(nameof(packagePart));
        if (string.IsNullOrEmpty(authorName)) 
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(authorName));

        authorInitials ??= GetInitialsFromName(authorName);
        dateTime ??= DateTime.Now;

        SetOwner(document, packagePart, false);
        Xml = Resource.CommentElement(authorName, authorInitials, dateTime.Value);
    }

    /// <summary>
    /// Turn an author name into initials by taking the first letter of each word.
    /// </summary>
    /// <param name="authorName"></param>
    /// <returns></returns>
    private static string GetInitialsFromName(string authorName) 
        => string.Join("", authorName.Split(' ')
            .Where(w => w.Length > 0)
            .Select(w => char.ToUpper(w[0]).ToString()));

    /// <summary>
    /// Constructor when pulling from an existing document
    /// </summary>
    /// <param name="document"></param>
    /// <param name="packagePart"></param>
    /// <param name="xml"></param>
    internal Comment(Document document, PackagePart packagePart, XElement xml) : base(xml)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (packagePart == null) throw new ArgumentNullException(nameof(packagePart));
        SetOwner(document, packagePart, false);
    }

    /// <summary>
    /// Override for Object.ToString().
    /// </summary>
    /// <returns></returns>
    public override string ToString() => $"{Id}: {Date?.ToString("s")} by {AuthorName}";

    /// <summary>
    /// Returns a list of all Paragraphs inside this container.
    /// </summary>
    public IEnumerable<Paragraph> Paragraphs
    {
        get
        {
            int current = 0;
            foreach (var e in Xml.Elements(Name.Paragraph))
            {
                yield return DocumentHelpers.WrapParagraphElement(e, Document, PackagePart, ref current);
            }
        }
    }

    /// <summary>
    /// Removes paragraph at specified position
    /// </summary>
    /// <param name="index">Index of paragraph to remove</param>
    /// <returns>True if removed</returns>
    public bool RemoveParagraph(int index)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index));

        var e = Xml.Descendants(Name.Paragraph).Skip(index).FirstOrDefault();
        e?.Remove();
        return e != null;
    }

    /// <summary>
    /// Removes paragraph
    /// </summary>
    /// <param name="paragraph">FirstParagraph to remove</param>
    /// <returns>True if removed</returns>
    public bool RemoveParagraph(Paragraph paragraph)
    {
        if (paragraph.Xml.Parent == this.Xml)
        {
            paragraph.Xml.Remove();
            return true;
        }
        return false;
    }

    /// <summary>
    /// Add a paragraph from some text to the comment.
    /// </summary>
    /// <param name="text">Text to add</param>
    public void AddParagraph(string text) => AddParagraph(new Paragraph(text));

    /// <summary>
    /// Add a paragraph to this comment.
    /// </summary>
    /// <param name="paragraph">Paragraph to add</param>
    public void AddParagraph(Paragraph paragraph)
    {
        if (paragraph == null) 
            throw new ArgumentNullException(nameof(paragraph));

        if (paragraph.Xml.InDom())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

        // Ensure we have the proper styles
        paragraph.Properties.StyleName = ParagraphStyle;
        foreach (var run in paragraph.Runs)
        {
            run.StyleName = RunStyle;
            if (run.Xml.Element(Namespace.Main + "annotationRef") == null)
                run.Xml.Add(new XElement(Namespace.Main + "annotationRef"));
        }

        // Add an ID if it's missing.
        if (paragraph.Xml.Attribute(Name.ParagraphId) == null)
        {
            paragraph.Xml.SetAttributeValue(Name.ParagraphId, DocumentHelpers.GenerateHexId());
        }

        // Add it to this document
        Xml.Add(paragraph.Xml);
    }

    /// <summary>
    /// Determines equality for comments.
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Comment? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a comment
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Comment);

    /// <summary>
    /// Returns hashcode for this comment
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();

}