using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Base container object - this is used to represent all elements that can contain child elements
/// </summary>
public abstract class BlockContainer : DocXElement, IContainer
{
    /// <summary>
    /// Constructor
    /// </summary>
    protected BlockContainer()
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="xml">XML element that represents this block.</param>
    protected BlockContainer(XElement xml) : base(xml)
    {
    }

    /// <summary>
    /// Enumerates all blocks in this block container.
    /// </summary>
    public IEnumerable<Block> Blocks
    {
        get
        {
            int current = 0;
            foreach (var e in Xml.Elements())
            {
                var block = WrapElementBlock(this, e, ref current);
                if (block != null) yield return block;
            }
        }
    }

    /// <summary>
    /// Enumerates the paragraphs inside this container.
    /// </summary>
    public IEnumerable<Paragraph> Paragraphs
    {
        get
        {
            int current = 0;
            foreach (var e in Xml.Elements(Name.Paragraph))
            {
                yield return DocumentHelpers.WrapParagraphElement(e, SafeDocument, SafePackagePart, ref current);
            }
        }
    }

    /// <summary>
    /// Removes paragraph at specified position
    /// </summary>
    /// <param name="index">Index of paragraph to remove</param>
    /// <returns>True if removed</returns>
    public bool RemoveAt(int index)
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
    public bool Remove(Paragraph paragraph)
    {
        if (paragraph.Container == this)
        {
            paragraph.Xml.Remove();
            return true;
        }
        return false;
    }

    /// <summary>
    /// Returns all the sections associated with this container.
    /// </summary>
    public IEnumerable<Section> Sections
    {
        get
        {
            if (!InDocument) yield break;

            // Return all the section dividers.
            foreach (var sp in Xml.Descendants(Name.ParagraphProperties)
                                          .Descendants(Name.SectionProperties)
                                          .Select(xe => xe.FindParent(Name.Paragraph))
                                          .OmitNull())
                yield return new Section(Document, PackagePart, sp);

            // Return the final section if this is the mainDoc.
            if (Xml.Element(Name.SectionProperties) != null)
                yield return new Section(Document, PackagePart, Xml);
        }
    }

    /// <summary>
    /// Retrieve a list of all Table objects in the document
    /// </summary>
    public IEnumerable<Table> Tables => Xml.Descendants(Name.Table)
        .Select(t => new Table(Document, PackagePart, t));

    /// <summary>
    /// Retrieve a list of all hyperlinks in the document
    /// </summary>
    public IEnumerable<Hyperlink> Hyperlinks => Paragraphs.SelectMany(p => p.Hyperlinks);

    /// <summary>
    /// Retrieve a list of all images (pictures) in the document
    /// </summary>
    public IEnumerable<Picture> Pictures => Paragraphs.SelectMany(p => p.Pictures);

    /// <summary>
    /// Replace matched text with a new value.
    /// </summary>
    /// <param name="findText">Text value to search for</param>
    /// <param name="replaceText">Replacement value</param>
    /// <param name="comparisonType">How to compare text</param>
    public bool FindReplace(string findText, string? replaceText, StringComparison comparisonType)
    {
        if (string.IsNullOrEmpty(findText)) throw new ArgumentNullException(nameof(findText));

        bool found = false;
        Sections.SelectMany(s => s.Headers).SelectMany(header => header.Paragraphs)
            .Union(Paragraphs)
            .Union(Sections.SelectMany(s => s.Footers).SelectMany(footer => footer.Paragraphs))
            .ToList()
            .ForEach(p => found|=p.FindReplace(findText, replaceText, comparisonType));
        return found;
    }

    /// <summary>
    /// Insert a text block at a specific bookmark
    /// </summary>
    /// <param name="bookmarkName">Bookmark name</param>
    /// <param name="toInsert">Text to insert</param>
    public bool InsertTextAtBookmark(string bookmarkName, string toInsert)
    {
        if (string.IsNullOrWhiteSpace(bookmarkName))
        {
            throw new ArgumentException("bookmark cannot be null or empty", nameof(bookmarkName));
        }

        // Try headers first
        if (Sections.SelectMany(s => s.Headers)
            .SelectMany(header => header.Paragraphs)
            .Any(paragraph => paragraph.InsertTextAtBookmark(bookmarkName, toInsert)))
        {
            return true;
        }

        // Body
        if (Paragraphs.Any(paragraph => paragraph.InsertTextAtBookmark(bookmarkName, toInsert)))
        {
            return true;
        }

        // Footers
        return Sections.SelectMany(s => s.Footers)
            .SelectMany(header => header.Paragraphs)
            .Any(paragraph => paragraph.InsertTextAtBookmark(bookmarkName, toInsert));
    }

    /// <summary>
    /// Insert a paragraph into this container at a specific index
    /// </summary>
    /// <param name="index">Character index to insert into</param>
    /// <param name="paragraph">FirstParagraph to insert</param>
    /// <returns>Inserted paragraph</returns>
    public Paragraph Insert(int index, Paragraph paragraph)
    {
        if (paragraph.Xml.InDom())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));
        if (!InDocument)
            throw new InvalidOperationException("Must be part of document structure.");

        var targetParagraph = Document.FindParagraphByIndex(index);
        if (targetParagraph == null)
        {
            AddElementToContainer(paragraph.Xml);
        }
        else
        {
            var (leftElement, rightElement) = targetParagraph.Split(index - targetParagraph.StartIndex!.Value);
            targetParagraph.Xml.ReplaceWith(leftElement, paragraph.Xml, rightElement);
        }

        return OnAddParagraph(paragraph);
    }

    /// <summary>
    /// Add a paragraph at the end of the container
    /// </summary>
    public Paragraph AddParagraph(Paragraph paragraph)
    {
        if (paragraph == null)
            throw new ArgumentNullException(nameof(paragraph));

        if (paragraph.Xml.InDom())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

        AddElementToContainer(paragraph.Xml);
        return OnAddParagraph(paragraph);
    }

    /// <summary>
    /// Create a set of paragraphs from a series of strings. Each string will add a new paragraph
    /// to the document.
    /// </summary>
    /// <param name="paragraphs">Other paragraphs</param>
    /// <returns>Last paragraph added, or null if no paragraphs were added</returns>
    public Paragraph? AddRange(IEnumerable<string> paragraphs)
    {
        if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));
        return AddRange(paragraphs.Select(t => new Paragraph(t)));
    }

    /// <summary>
    /// Create a set of paragraphs from a series of strings. Each string will add a new paragraph
    /// to the document.
    /// </summary>
    /// <param name="paragraphs">Paragraphs to add</param>
    /// <returns>Last paragraph added, or null if no paragraphs were added</returns>
    public Paragraph? AddRange(IEnumerable<Paragraph> paragraphs)
    {
        if (paragraphs == null) throw new ArgumentNullException(nameof(paragraphs));

        Paragraph? lastParagraph = null;
        foreach (var p in paragraphs)
        {
            lastParagraph = Add(p);
        }

        return lastParagraph;
    }

    /// <summary>
    /// This method is called when a new paragraph is added to this container.
    /// </summary>
    /// <param name="paragraph">New paragraph</param>
    /// <returns>Added paragraph</returns>
    internal Paragraph OnAddParagraph(Paragraph paragraph)
    {
        if (string.IsNullOrEmpty(paragraph.Id))
        {
            paragraph.Xml.SetAttributeValue(Name.ParagraphId, DocumentHelpers.GenerateHexId());
        }

        paragraph.SetOwner(Document, PackagePart, true);
        paragraph.StartIndex = Paragraphs.Single(p => p.Id == paragraph.Id).StartIndex!.Value;

        return paragraph;
    }

    /// <summary>
    /// Adds a new paragraph into the document structure
    /// </summary>
    /// <param name="xml"></param>
    protected virtual XElement AddElementToContainer(XElement xml)
    {
        // On paragraphs, add an ID if it's missing.
        if (xml.Name == Name.Paragraph
            && xml.Attribute(Name.ParagraphId) == null)
        {
            xml.SetAttributeValue(Name.ParagraphId, DocumentHelpers.GenerateHexId());
        }

        // If this is the body document, then add the paragraph just before the
        // body section -- this must always be the final thing in the document.
        var sectPr = Xml.Elements(Name.SectionProperties).SingleOrDefault();
        if (sectPr != null)
        {
            sectPr.AddBeforeSelf(xml);
        }
        // Otherwise, we can just add to the end of the XML block.
        else
        {
            Xml.Add(xml);
        }

        return xml;
    }

    /// <summary>
    /// Add a new section to the container
    /// </summary>
    /// <param name="breakType">The type of section break to insert</param>
    public virtual Section AddSection(SectionBreakType breakType)
    {
        if (Document == null || PackagePart == null)
            throw new InvalidOperationException("Must be part of document structure.");

        var xml = AddElementToContainer(
            new XElement(Name.Paragraph,
                new XAttribute(Name.ParagraphId, DocumentHelpers.GenerateHexId()),
                new XElement(Name.ParagraphProperties,
                    new XElement(Name.SectionProperties,
                        new XElement(Namespace.Main + "type", new XAttribute(Name.MainVal, breakType.GetEnumName())),
                        new XElement(Namespace.Main + "cols", new XAttribute(Namespace.Main + "space", 720)),
                        new XElement(Namespace.Main + "docGrid", new XAttribute(Namespace.Main+"linePitch", 360))))));

        return new Section(Document, PackagePart, xml);
    }

    /// <summary>
    /// Add a new page break to the container
    /// </summary>
    public virtual void AddPageBreak()
    {
        AddElementToContainer(new XElement(Name.Paragraph,
            new XAttribute(Name.ParagraphId, DocumentHelpers.GenerateHexId()),
            new XElement(Name.Run,
                new XElement(Name.Break,
                    new XAttribute(Name.Type, "page")))));
    }

    /// <summary>
    /// Add a paragraph with the given text to the end of the container
    /// </summary>
    /// <param name="paragraph">Text to add</param>
    /// <returns></returns>
    public Paragraph Add(Paragraph paragraph)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        if (paragraph.InDocument)
            paragraph = new Paragraph(Document, PackagePart, paragraph.Xml.Normalize(), 0);

        AddElementToContainer(paragraph.Xml);
        return OnAddParagraph(paragraph);
    }

    /// <summary>
    /// Add a new table to the end of the container
    /// </summary>
    /// <param name="table">Table to add</param>
    /// <returns>The table now associated with the document.</returns>
    public Table Add(Table table)
    {
        if (table.Xml.InDom())
            throw new ArgumentException("Cannot add table multiple times.", nameof(table));

        // The table will be added to the end of the document. If the
        // prior element is _also_ a table, then we need to insert a blank
        // paragraph first - otherwise Word will merge the tables. It's
        // acceptable to have a table be the only element.
        var lastElementInDoc = Xml
            .Elements()
            .LastOrDefault(e => e.Name != Name.SectionProperties);
        if (lastElementInDoc?.Name == Name.Table)
        {
            this.AddParagraph();
        }

        AddElementToContainer(table.Xml);

        // Set the document and package
        table.SetOwner(Document, PackagePart, true);

        return table;
    }

    /// <summary>
    /// Insert a Table into this document.
    /// </summary>
    /// <param name="index">The index to insert this Table at.</param>
    /// <param name="table">The Table to insert.</param>
    /// <returns>The Table now associated with this document.</returns>
    public Table Insert(int index, Table table)
    {
        if (table.Xml.InDom())
            throw new ArgumentException("Cannot add table multiple times.", nameof(table));
        if (Document == null)
            throw new InvalidOperationException("Must be part of document structure.");

        table.SetOwner(Document, PackagePart, true);

        var firstParagraph = Document.FindParagraphByIndex(index);
        if (firstParagraph != null)
        {
            var (leftElement, rightElement) = firstParagraph.Split(index - firstParagraph.StartIndex!.Value);
            firstParagraph.Xml.ReplaceWith(leftElement, table.Xml, rightElement);
        }

        return table;
    }

    /// <summary>
    /// Called when the document owner is changed.
    /// </summary>
    protected override void OnAddToDocument()
    {
        foreach (var paragraph in Paragraphs)
        {
            paragraph.SetOwner(Document, PackagePart, true);
        }
    }

    /// <summary>
    /// Helper to create a block from an element in the document.
    /// </summary>
    /// <param name="blockContainer">Owning container</param>
    /// <param name="e">XML element</param>
    /// <param name="current">Current text position for paragraph tracking</param>
    /// <returns>Block wrapper</returns>
    private static Block? WrapElementBlock(BlockContainer blockContainer, XElement e, ref int current)
    {
        if (blockContainer == null) throw new ArgumentNullException(nameof(blockContainer));
        if (e == null) throw new ArgumentNullException(nameof(e));

        return e.Name == Name.Paragraph
            ? DocumentHelpers.WrapParagraphElement(e, blockContainer.Document, blockContainer.PackagePart, ref current)
            : e.Name == Name.Table
                ? new Table(blockContainer.Document, blockContainer.PackagePart, e)
                : e.Name != Name.SectionProperties
                    ? new UnknownBlock(blockContainer.Document, blockContainer.PackagePart, e)
                    : null;
    }
}