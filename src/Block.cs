using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a Paragraph or Table in the document.
/// </summary>
public abstract class Block : DocXElement
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="xml">XML for this block</param>
    protected Block(XElement xml) : base(xml)
    {
    }

    /// <summary>
    /// Retrieves the container owner (TableCell, Document, Header/Footer).
    /// </summary>
    internal BlockContainer? Container
    {
        get
        {
            if (!Xml.HasParent()) return null;
            if (Xml.Parent?.Name == Name.Body) return this.Document;
            return Xml.Parent != null
                ? WrapElementBlockContainer(this.Document, this.PackagePart, Xml.Parent)
                : null;
        }
    }


    /// <summary>
    /// Add a page break after the current element.
    /// </summary>
    public void AddPageBreak() =>
        Xml.AddAfterSelf(CreatePageBreakElement);

    /// <summary>
    /// Insert a page break before the current element.
    /// </summary>
    public void InsertPageBreakBefore() =>
        Xml.AddBeforeSelf(CreatePageBreakElement);

    /// <summary>
    /// Insert a paragraph before this one in the document.
    /// </summary>
    /// <param name="paragraph">Paragraph to add.</param>
    /// <returns>Added paragraph</returns>
    public Paragraph InsertBefore(Paragraph paragraph)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        if (!InDocument)
            throw new InvalidOperationException("Can only append to paragraphs in an existing document structure.");
        if (paragraph.Xml.HasParent())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

        Xml.AddBeforeSelf(paragraph.Xml);
        Document.OnAddParagraph(paragraph);
        return paragraph;
    }

    /// <summary>
    /// Add a new paragraph after the current element.
    /// </summary>
    /// <param name="paragraph">Paragraph to add</param>
    public Paragraph InsertAfter(Paragraph paragraph)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        if (!InDocument)
            throw new InvalidOperationException("Cannot add paragraphs to unowned paragraphs - must be part of a document structure.");
        if (paragraph.Xml.HasParent())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

        // If this element is a paragraph and it currently has a table after it, then add the given paragraph
        // after the table.
        if (this is Paragraph p && p.Table != null)
        {
            p.Table.Xml.AddAfterSelf(paragraph.Xml);
        }
        // Otherwise, add the paragraph after this element.
        else
        {
            Xml.AddAfterSelf(paragraph.Xml);
        }

        // Set the id + owner information.
        Container!.OnAddParagraph(paragraph);

        return paragraph;
    }

    /// <summary>
    /// Wraps a block container from the document. These are elements
    /// which contain other elements.
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package part</param>
    /// <param name="e">Element</param>
    /// <returns></returns>
    private static BlockContainer? WrapElementBlockContainer(Document? document, PackagePart? packagePart, XElement e)
    {
        if (e.Name == Name.TableCell)
        {
            if (e.Parent?.Parent != null)
            {
                var rowXml = e.Parent;
                var tableXml = rowXml.Parent;
                var table = new Table(document, packagePart, tableXml);
                var row = new TableRow(table, rowXml);
                return new TableCell(row, e);
            }
        }

        if (e.Name == Name.Body)
        {
            return document;
        }

        if (e.Name.LocalName == "hdr" && document != null)
        {
            return document.Sections.SelectMany(s => s.Headers)
                .SingleOrDefault(h => h.Xml == e);
        }

        if (e.Name.LocalName == "ftr" && document != null)
        {
            return document.Sections.SelectMany(s => s.Footers)
                .SingleOrDefault(h => h.Xml == e);
        }

        throw new Exception($"Unrecognized container type {e.Name}");
    }

    /// <summary>
    /// Create a page break paragraph element
    /// </summary>
    /// <returns>Page break element</returns>
    private static XElement CreatePageBreakElement => new(Name.Paragraph,
        new XAttribute(Name.ParagraphId, DocumentHelpers.GenerateHexId()),
        new XElement(Name.Run, new XElement(Namespace.Main + RunTextType.LineBreak,
            new XAttribute(Namespace.Main + "type", "page"))));
}