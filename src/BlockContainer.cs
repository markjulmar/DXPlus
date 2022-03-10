using DXPlus.Helpers;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml.Linq;

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
                var block = HelperFunctions.WrapElementBlock(this, e, ref current);
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
                yield return HelperFunctions.WrapParagraphElement(e, SafeDocument, SafePackagePart, ref current);
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
            foreach (var p in Paragraphs.Where(p => p.IsSectionParagraph))
                yield return new Section(Document, PackagePart, p.Xml);

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
    /// TODO: add complex search object to simplify signature
    /// </summary>
    /// <param name="searchValue">Text value to search for</param>
    /// <param name="newValue">Replacement value</param>
    /// <param name="options">Regex options</param>
    /// <param name="newFormatting">New formatting to apply</param>
    /// <param name="matchFormatting">Formatting to match</param>
    /// <param name="formattingOptions">Match formatting options</param>
    /// <param name="escapeRegEx">True to escape Regex expression</param>
    /// <param name="useRegExSubstitutions">True to use RegEx in substitution</param>
    public void ReplaceText(string searchValue, string newValue,
        RegexOptions options = RegexOptions.None, Formatting? newFormatting = null, Formatting? matchFormatting = null,
        MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch,
        bool escapeRegEx = true, bool useRegExSubstitutions = false)
    {
        if (string.IsNullOrEmpty(searchValue))
        {
            throw new ArgumentNullException(nameof(searchValue));
        }

        if (newValue == null)
        {
            throw new ArgumentNullException(nameof(newValue));
        }

        // ReplaceText in Headers of all sections.
        foreach (Paragraph paragraph in Sections.SelectMany(s => s.Headers)
                     .SelectMany(header => header.Paragraphs))
        {
            paragraph.ReplaceText(searchValue, newValue, options,
                newFormatting, matchFormatting, formattingOptions,
                escapeRegEx, useRegExSubstitutions);
        }

        // ReplaceText int main body of document.
        foreach (Paragraph paragraph in Paragraphs)
        {
            paragraph.ReplaceText(searchValue, newValue, options,
                newFormatting, matchFormatting, formattingOptions,
                escapeRegEx, useRegExSubstitutions);
        }

        // ReplaceText in Footers of the document.
        foreach (Paragraph paragraph in Sections.SelectMany(s => s.Footers)
                     .SelectMany(footer => footer.Paragraphs))
        {
            paragraph.ReplaceText(searchValue, newValue, options,
                newFormatting, matchFormatting, formattingOptions,
                escapeRegEx, useRegExSubstitutions);
        }
    }

    /// <summary>
    /// Insert a text block at a specific bookmark
    /// </summary>
    /// <param name="bookmarkName">Bookmark name</param>
    /// <param name="toInsert">Text to insert</param>
    public bool InsertAtBookmark(string bookmarkName, string toInsert)
    {
        if (string.IsNullOrWhiteSpace(bookmarkName))
        {
            throw new ArgumentException("bookmark cannot be null or empty", nameof(bookmarkName));
        }

        // Try headers first
        if (Sections.SelectMany(s => s.Headers)
            .SelectMany(header => header.Paragraphs)
            .Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert)))
        {
            return true;
        }

        // Body
        if (Paragraphs.Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert)))
        {
            return true;
        }

        // Footers
        return Sections.SelectMany(s => s.Footers)
            .SelectMany(header => header.Paragraphs)
            .Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert));
    }

    /// <summary>
    /// Insert a paragraph into this container at a specific index
    /// </summary>
    /// <param name="index">Character index to insert into</param>
    /// <param name="paragraph">FirstParagraph to insert</param>
    /// <returns>Inserted paragraph</returns>
    public Paragraph InsertParagraph(int index, Paragraph paragraph)
    {
        if (paragraph.Xml.InDom())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));
        if (Document == null)
            throw new InvalidOperationException("Must be part of document structure.");

        var insertPos = Document.FindParagraphByIndex(index);
        if (insertPos == null)
        {
            AddElementToContainer(paragraph.Xml);
        }
        else
        {
            var split = SplitParagraph(insertPos, index - insertPos.StartIndex);
            insertPos.Xml.ReplaceWith(split[0], paragraph.Xml, split[1]);
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
    /// This method is called when a new paragraph is added to this container.
    /// </summary>
    /// <param name="paragraph">New paragraph</param>
    /// <returns>Added paragraph</returns>
    internal Paragraph OnAddParagraph(Paragraph paragraph)
    {
        if (string.IsNullOrEmpty(paragraph.Id))
        {
            paragraph.Xml.SetAttributeValue(Name.ParagraphId, HelperFunctions.GenerateHexId());
        }

        InsertMissingStyles(paragraph);

        paragraph.SetOwner(Document, PackagePart, true);
        paragraph.SetStartIndex(Paragraphs.Single(p => p.Id == paragraph.Id).StartIndex);

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
            xml.SetAttributeValue(Name.ParagraphId, HelperFunctions.GenerateHexId());
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
    /// Insert any missing styles associated with the passed paragraph
    /// </summary>
    /// <param name="paragraph"></param>
    private void InsertMissingStyles(Paragraph paragraph)
    {
        if (Document?.Package == null)
            return;

        // Make sure the document has all the styles associated to the
        // paragraph we are inserting.
        if (paragraph.Styles.Count > 0)
        {
            Uri stylePackage = Relations.Styles.Uri;
            XDocument styleDoc;
            PackagePart stylePackagePart;
            if (!Document.Package.PartExists(stylePackage))
            {
                stylePackagePart = Document.Package.CreatePart(stylePackage, Relations.Styles.ContentType, CompressionOption.Maximum);
                styleDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(Namespace.Main + "styles"));
                stylePackagePart.Save(styleDoc);
            }
            else
            {
                stylePackagePart = Document.Package.GetPart(stylePackage);
                styleDoc = stylePackagePart.Load();
            }

            // Get all the styleId values from the current style
            XElement styles = styleDoc.GetOrAddElement(Namespace.Main + "styles");
            List<string> ids = styles.Descendants(Namespace.Main + "style")
                .Select(e => e.AttributeValue(Namespace.Main + "styleId", null))
                .Where(v => v != null)
                .Cast<string>()
                .ToList();

            // Go through the new paragraph and make sure all the styles are present
            foreach (var style in paragraph.Styles.Where(s => 
                    !ids.Contains(s.AttributeValue(Namespace.Main + "styleId")??"")))
            {
                styles.Add(style);
            }

            // Save back to the package
            stylePackagePart.Save(styleDoc);
        }
    }

    /// <summary>
    /// Insert a new paragraph using the passed text.
    /// </summary>
    /// <param name="index">Index to insert into</param>
    /// <param name="text">Text for new paragraph</param>
    /// <param name="formatting">Formatting for new paragraph</param>
    /// <returns></returns>
    public Paragraph InsertParagraph(int index, string text, Formatting? formatting)
    {
        if (Document == null)
            throw new InvalidOperationException("Must be part of document structure.");

        var newParagraph = new Paragraph(Document, PackagePart, Paragraph.Create(text, formatting), index);
        var firstPar = Document.FindParagraphByIndex(index);
        if (firstPar != null)
        {
            int splitIndex = index - firstPar.StartIndex;
            if (splitIndex <= 0)
            {
                firstPar.Xml.ReplaceWith(newParagraph.Xml, firstPar.Xml);
            }
            else
            {
                var splitParagraph = SplitParagraph(firstPar, splitIndex);
                firstPar.Xml.ReplaceWith(splitParagraph[0], newParagraph.Xml, splitParagraph[1]);
            }
        }
        else
        {
            AddElementToContainer(newParagraph.Xml);
        }

        return OnAddParagraph(newParagraph);
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
                new XAttribute(Name.ParagraphId, HelperFunctions.GenerateHexId()),
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
            new XAttribute(Name.ParagraphId, HelperFunctions.GenerateHexId()),
            new XElement(Name.Run,
                new XElement(Name.Break,
                    new XAttribute(Name.Type, "page")))));
    }

    /// <summary>
    /// Add a paragraph with the given text to the end of the container
    /// </summary>
    /// <param name="text">Text to add</param>
    /// <param name="formatting">Formatting to use</param>
    /// <returns></returns>
    public Paragraph AddParagraph(string text, Formatting? formatting)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));

        var paragraph = Paragraph.Create(text, formatting);
        AddElementToContainer(paragraph);
        return OnAddParagraph(new Paragraph(Document, PackagePart, paragraph, 0));
    }

    /// <summary>
    /// Add a new table to the end of the container
    /// </summary>
    /// <param name="table">Table to add</param>
    /// <returns>The table now associated with the document.</returns>
    public Table AddTable(Table table)
    {
        if (table.Xml.InDom())
            throw new ArgumentException("Cannot add table multiple times.", nameof(table));

        table.SetOwner(Document, PackagePart, true);

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

        return table;
    }

    /// <summary>
    /// Insert a Table into this document.
    /// </summary>
    /// <param name="index">The index to insert this Table at.</param>
    /// <param name="table">The Table to insert.</param>
    /// <returns>The Table now associated with this document.</returns>
    public Table InsertTable(int index, Table table)
    {
        if (table.Xml.InDom())
            throw new ArgumentException("Cannot add table multiple times.", nameof(table));
        if (Document == null)
            throw new InvalidOperationException("Must be part of document structure.");

        table.SetOwner(Document, PackagePart, true);

        var firstParagraph = Document.FindParagraphByIndex(index);
        if (firstParagraph != null)
        {
            var split = SplitParagraph(firstParagraph, index - firstParagraph.StartIndex);
            firstParagraph.Xml.ReplaceWith(split[0], table.Xml, split[1]);
        }

        return table;
    }

    /// <summary>
    /// Called when the document owner is changed.
    /// </summary>
    protected override void OnAddToDocument()
    {
        // Make sure we have all the styles in our document owner.
        Paragraphs.ToList().ForEach(InsertMissingStyles);
    }
        
    /// <summary>
    /// Split a paragraph at a specific index
    /// </summary>
    /// <param name="paragraph">FirstParagraph to split</param>
    /// <param name="index">Character index to split at</param>
    /// <returns>Left/Right split</returns>
    /// <exception cref="ArgumentNullException"></exception>
    private static XElement?[] SplitParagraph(Paragraph paragraph, int index)
    {
        if (paragraph == null)
            throw new ArgumentNullException(nameof(paragraph));
            
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index));

        XElement?[] split = {null, null};

        var r = paragraph.FindRunAffectedByEdit(EditType.Insert, index);
        if (r == null)
            return split;

        XElement? before, after;
        switch (r.Xml.Parent?.Name.LocalName)
        {
            case "ins":
                split = paragraph.SplitEdit(r.Xml.Parent, index, EditType.Insert);
                before = new XElement(paragraph.Xml.Name, paragraph.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(paragraph.Xml.Name, paragraph.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
                break;

            case "del":
                split = paragraph.SplitEdit(r.Xml.Parent, index, EditType.Delete);
                before = new XElement(paragraph.Xml.Name, paragraph.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(paragraph.Xml.Name, paragraph.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
                break;

            default:
                split = r.SplitAtIndex(index);
                before = new XElement(paragraph.Xml.Name, paragraph.Xml.Attributes(), r.Xml.ElementsBeforeSelf(), split[0]);
                after = new XElement(paragraph.Xml.Name, paragraph.Xml.Attributes(), split[1], r.Xml.ElementsAfterSelf());
                break;
        }

        if (!before.Elements().Any())
        {
            before = null;
        }

        if (!after.Elements().Any())
        {
            after = null;
        }

        return new[] { before, after };
    }
}