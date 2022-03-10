using DXPlus.Helpers;
using System.Xml.Linq;

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
            if (!Xml.InDom()) return null;
            if (Xml.Parent?.Name == Name.Body) return this.Document;
            return Xml.Parent != null
                ? HelperFunctions.WrapElementBlockContainer(this.Document, this.PackagePart, Xml.Parent)
                : null;
        }
    }


    /// <summary>
    /// Add a page break after the current element.
    /// </summary>
    public void AddPageBreak() =>
        Xml.AddAfterSelf(HelperFunctions.PageBreak());

    /// <summary>
    /// Insert a page break before the current element.
    /// </summary>
    public void InsertPageBreakBefore() =>
        Xml.AddBeforeSelf(HelperFunctions.PageBreak());

    /// <summary>
    /// Add an empty paragraph after the current element.
    /// </summary>
    /// <returns>Created empty paragraph</returns>
    public Paragraph AddParagraph() => AddParagraph(string.Empty, null);

    /// <summary>
    /// Add a new paragraph after the current element.
    /// </summary>
    /// <param name="paragraph">FirstParagraph to insert</param>
    public Paragraph AddParagraph(Paragraph paragraph)
    {
        if (!Xml.InDom())
            throw new InvalidOperationException("Cannot add paragraphs to unowned paragraphs - must be part of a document structure.");
        if (paragraph == null) 
            throw new ArgumentNullException(nameof(paragraph));
        if (paragraph.Xml.InDom())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

        // If this element is a paragraph and it currently
        // has a table after it, then add the given paragraph
        // after the table.
        if (this is Paragraph p && p.Table != null)
        {
            p.Table.Xml.AddAfterSelf(paragraph.Xml);
        }
        // Otherwise, just add the paragraph.
        else Xml.AddAfterSelf(paragraph.Xml);

        // Determine the starting index using the container owner.
        // TODO: remove
        paragraph.SetStartIndex(Container!.Paragraphs.Single(p => p.Id == paragraph.Id).StartIndex);

        return paragraph;
    }

    /// <summary>
    /// Add a paragraph after the current element using the passed text
    /// </summary>
    /// <param name="text">Text for new paragraph</param>
    /// <param name="formatting">Formatting for the paragraph</param>
    /// <returns>Newly created paragraph</returns>
    public Paragraph AddParagraph(string text, Formatting? formatting)
    {
        var paragraph = new Paragraph(Document, PackagePart, Paragraph.Create(text, formatting), -1);
        AddParagraph(paragraph);
        return paragraph;
    }

    /// <summary>
    /// Insert a paragraph before the current element
    /// </summary>
    /// <param name="paragraph"></param>
    public void InsertParagraphBefore(Paragraph paragraph)
    {
        if (paragraph == null) 
            throw new ArgumentNullException(nameof(paragraph));
        if (paragraph.Xml.InDom())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

        Xml.AddBeforeSelf(paragraph.Xml);

        // TODO: remove
        paragraph.SetStartIndex(Container!.Paragraphs.Single(p => p.Id == paragraph.Id).StartIndex);
    }

    /// <summary>
    /// Insert a paragraph before the current element
    /// </summary>
    /// <param name="text">Text to use for new paragraph</param>
    /// <param name="formatting">Formatting to use</param>
    /// <returns></returns>
    public Paragraph InsertParagraphBefore(string text, Formatting? formatting)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));
        var paragraph = new Paragraph(Document, PackagePart, Paragraph.Create(text, formatting), -1);
        InsertParagraphBefore(paragraph);
        return paragraph;
    }

    /// <summary>
    /// Insert a table before this element
    /// </summary>
    /// <param name="table"></param>
    public void InsertTableBefore(Table table)
    {
        if (table == null) 
            throw new ArgumentNullException(nameof(table));
        if (table.Xml.InDom())
            throw new ArgumentException("Cannot add table multiple times.", nameof(table));

        table.SetOwner(Document, PackagePart, true);
        Xml.AddBeforeSelf(table.Xml);
    }
}