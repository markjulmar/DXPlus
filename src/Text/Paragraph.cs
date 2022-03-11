﻿using DXPlus.Helpers;
using System.Diagnostics;
using System.Drawing;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a document paragraph.
/// </summary>
[DebuggerDisplay("({StartIndex}-{EndIndex}) - {Text}")]
public class Paragraph : Block, IEquatable<Paragraph>
{
    private Table? tableAfterParagraph;
    private readonly List<Hyperlink> unownedHyperlinks = new();

    /// <summary>
    /// Unique id for this paragraph
    /// </summary>
    public string? Id
    {
        get => Xml.AttributeValue(Name.ParagraphId, null);
        private init => Xml.SetAttributeValue(Name.ParagraphId, string.IsNullOrEmpty(value) ? null : value);
    }

    /// <summary>
    /// Text runs (r) that make up this paragraph
    /// </summary>
    public IEnumerable<Run> Runs
    {
        get
        {
            int start = 0;
            // Only look at the localName so we capture Math.r and Main.r
            foreach (var runXml in Xml.Descendants().Where(x => x.Name.LocalName == Name.Run.LocalName))
            {
                var run = new Run(SafeDocument, SafePackagePart, runXml, start);
                yield return run;
                start = run.EndIndex;
            }
        }
    }

    /// <summary>
    /// Styles in this paragraph
    /// TODO: implement
    /// </summary>
    internal List<XElement> Styles { get; } = new();

    /// <summary>
    /// Starting index for this paragraph
    /// TODO: remove
    /// </summary>
    internal int StartIndex { get; private set; }

    /// <summary>
    /// End index for this paragraph
    /// TODO: remove
    /// </summary>
    internal int EndIndex { get; private set; }

    /// <summary>
    /// Public constructor for the paragraph
    /// </summary>
    public Paragraph() 
        : this(null, null, Create(string.Empty, null), 0)
    {
    }

    /// <summary>
    /// Public constructor for the paragraph
    /// </summary>
    /// <param name="text"></param>
    public Paragraph(string text) 
        : this (null, null, Create(text, null), 0)
    {
    }

    /// <summary>
    /// Public constructor for the paragraph
    /// </summary>
    /// <param name="text"></param>
    /// <param name="formatting"></param>
    public Paragraph(string text, Formatting formatting) 
        : this(null, null, Create(text, formatting), 0)
    {
    }

    /// <summary>
    /// Constructor for the paragraph
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package owner</param>
    /// <param name="xml">XML for the paragraph</param>
    /// <param name="startIndex">Starting position in the doc</param>
    internal Paragraph(Document? document, PackagePart? packagePart, XElement xml, int startIndex) : base(xml)
    {
        if (document != null)
        {
            SetOwner(document, packagePart, false);
        }

        StartIndex = startIndex;
        EndIndex = startIndex + HelperFunctions.GetTextLength(xml);
    }

    /// <summary>
    /// Attaches a comment to this paragraph.
    /// </summary>
    /// <param name="comment">Comment to add</param>
    public void AttachComment(Comment comment) => AttachComment(comment, Runs.First());

    /// <summary>
    /// Attach a comment to this Run
    /// </summary>
    /// <param name="comment">Comment</param>
    /// <param name="run">Text run</param>
    public void AttachComment(Comment comment, Run run) => AttachComment(comment, run, run);

    /// <summary>
    /// Attach a comment to this Run
    /// </summary>
    /// <param name="comment">Comment</param>
    /// <param name="runStart">Text run</param>
    /// <param name="runEnd">End run</param>
    public void AttachComment(Comment comment, Run runStart, Run runEnd)
    {
        if (comment == null) throw new ArgumentNullException(nameof(comment));
        if (runStart == null) throw new ArgumentNullException(nameof(runStart));
        if (runStart.Xml.Parent != Xml) throw new ArgumentException("Specified run not part of paragraph.", nameof(runStart));
        if (runEnd.Xml.Parent != Xml) throw new ArgumentException("Specified run not part of paragraph.", nameof(runEnd));

        if (Document == null)
            throw new Exception(
                "Cannot attached comments with unowned paragraph. Add the paragraph to a document first.");

        CommentManager.Attach(comment, runStart, runEnd);
    }

    /// <summary>
    /// Retrieve all bookmarks
    /// </summary>
    public BookmarkCollection Bookmarks => new(GetBookmarks());

    /// <summary>
    /// Comments tied to this paragraph
    /// </summary>
    public IEnumerable<CommentRange> Comments 
        => InDocument ? Document.CommentManager.GetCommentsForParagraph(this) : Enumerable.Empty<CommentRange>();

    /// <summary>
    /// The default run properties applied at the paragraph level
    /// </summary>
    public Formatting? DefaultFormatting
    {
        get
        {
            var e = Xml.Element(Name.ParagraphProperties, Name.RunProperties);
            return e != null ? new Formatting(e) : null;
        }

        set
        {
            var pPr = Xml.Element(Name.ParagraphProperties);
            if (pPr == null && value != null)
            {
                pPr = new XElement(Name.ParagraphProperties);
                Xml.AddFirst(pPr);
            }
            
            pPr?.GetRunProperties()?.Remove();
            if (value != null)
            {
                pPr!.AddFirst(value.Xml);
            }
        }
    }

    /// <summary>
    /// Apply the specified formatting to the paragraph or last findText run
    /// </summary>
    /// <param name="formatting">Formatting to apply</param>
    /// <returns>paragraph</returns>
    public Paragraph WithFormatting(Formatting formatting)
    {
        if (Runs.Any())
        {
            var runs = Runs.Reverse().ToList();
            var run = runs.Find(r => r.HasText) ?? runs[0];
            run.Properties = formatting;
        }
        else
        {
            DefaultFormatting = formatting;
        }

        return this;
    }

    /// <summary>
    /// Adds to the existing formatting for the paragraph and/or last run.
    /// </summary>
    /// <param name="formatting">Formatting to add</param>
    /// <returns>paragraph</returns>
    public Paragraph AddFormatting(Formatting formatting)
    {
        if (Runs.Any())
        {
            var runs = Runs.Reverse().ToList();
            var run = runs.Find(r => r.HasText) ?? runs[0];
            if (run.Properties == null) 
                run.Properties = formatting;
            else 
                run.Properties.Merge(formatting);
        }
        else
        {
            if (DefaultFormatting == null)
                DefaultFormatting = formatting;
            else
                DefaultFormatting.Merge(formatting);
        }
        return this;
    }

    /// <summary>
    /// Returns a list of DocProperty elements in this document.
    /// </summary>
    public IEnumerable<DocProperty> DocumentProperties
    {
        get
        {
            if (!InDocument) return Enumerable.Empty<DocProperty>();

            var properties = Xml.Descendants(Name.SimpleField)
                .Select(el => new DocProperty(Document, Document.PackagePart, el, null)).ToList();

            // Look for complex field insertions in the paragraph. These should always be in run elements and
            // have a start, name, sep, value section.
            foreach (var field in Xml.Descendants(Name.ComplexField).Where(e => e.AttributeValue(Namespace.Main + "fldCharType") == "begin").ToList())
            {
                // Start of a field. Walk down the tree looking for the name (instrText).
                var node = field.Parent;
                if (node == null || node.Name != Name.Run || node.Parent != Xml)
                    continue;

                XElement? nameNode = null;
                node = node.NextNode as XElement;
                while (node?.Parent == Xml)
                {
                    nameNode = node.Descendants(Namespace.Main + "instrText").SingleOrDefault();
                    if (nameNode != null)
                    {
                        nameNode = node; // want the run parent
                        break;
                    }
                }

                // Never found name.
                if (nameNode == null)
                    continue;

                // Look for the next [w:t] findText element.
                XElement? valueNode = null;
                node = nameNode.NextNode as XElement;
                while (node?.Parent == nameNode.Parent)
                {
                    valueNode = node?.Descendants(Name.Text).SingleOrDefault();
                    if (valueNode != null)
                    {
                        valueNode = node;
                        break;
                    }
                    node = node?.NextNode as XElement;
                }

                if (valueNode != null)
                {
                    properties.Add(new DocProperty(Document, Document.PackagePart, nameNode, valueNode));
                }
            }
            return properties;
        }
    }

    ///<summary>
    /// If a Table (tbl) follows this paragraph, then this property will have a reference to it.
    ///</summary>
    public Table? Table => !InDocument 
            ? tableAfterParagraph 
            : Xml.NextNode is XElement e && e.Name == Name.Table
                ? new Table(Document, PackagePart, e) : null;

    /// <summary>
    /// Retrieve the previous paragraph sibling.
    /// </summary>
    public Paragraph? PreviousParagraph
    {
        get
        {
            if (!InDocument) return null;
            var ps = Document.Paragraphs.ToList();
            int pos = ps.IndexOf(this);
            return pos > 0 ? ps[pos - 1] : null;
        }
    }

    /// <summary>
    /// Retrieve the next paragraph sibling.
    /// </summary>
    public Paragraph? NextParagraph
    {
        get
        {
            if (!InDocument) return null;
            var ps = Document.Paragraphs.ToList();
            int pos = ps.IndexOf(this);
            return pos >= 0 && pos < ps.Count-1 ? ps[pos+1] : null;
        }
    }

    /// <summary>
    /// True if this paragraph is a section divider.
    /// </summary>
    public bool IsSectionParagraph =>
        Xml.Element(Name.ParagraphProperties, Name.SectionProperties) != null;

    /// <summary>
    /// Returns the section this paragraph is associated with.
    /// </summary>
    public Section? Section => !InDocument 
        ? null 
        : Document.Sections.SingleOrDefault(s => s.Paragraphs.Contains(this));

    /// <summary>
    /// Append a paragraph after this one in the document.
    /// </summary>
    /// <param name="paragraph">Paragraph to add.</param>
    /// <returns>Added paragraph</returns>
    public Paragraph Append(Paragraph paragraph)
    {
        if (!InDocument)
            throw new InvalidOperationException("Can only append to paragraphs in an existing document structure.");
        if (paragraph.Xml.InDom())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

        Xml.AddAfterSelf(paragraph.Xml);
        Document.OnAddParagraph(paragraph);
        
        return paragraph;
    }

    /// <summary>
    /// Insert a paragraph before this one in the document.
    /// </summary>
    /// <param name="paragraph">Paragraph to add.</param>
    /// <returns>Added paragraph</returns>
    public Paragraph InsertBefore(Paragraph paragraph)
    {
        if (!InDocument)
            throw new InvalidOperationException("Can only append to paragraphs in an existing document structure.");
        if (paragraph.Xml.InDom())
            throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

        Xml.AddBeforeSelf(paragraph.Xml);
        Document.OnAddParagraph(paragraph);
        return paragraph;
    }

    /// <summary>
    /// Add a new table after this paragraph
    /// </summary>
    /// <param name="table">Table to add</param>
    public Paragraph Append(Table table)
    {
        if (table.Xml.InDom())
            throw new ArgumentException("Cannot add table multiple times.", nameof(table));

        if (Table != null)
            throw new Exception("Can only add one table after a paragraph. Must add paragraph separators between tables or they merge.");

        if (InDocument)
        {
            table.SetOwner(Document, PackagePart, true);
            Xml.AddAfterSelf(table.Xml);
        }
        else
        {
            tableAfterParagraph = table;
        }

        return this;
    }

    /// <summary>
    /// Properties applied to this paragraph
    /// </summary>
    public ParagraphProperties Properties
    {
        get
        {
            var pPr = Xml.Element(Name.ParagraphProperties);
            if (pPr == null)
                Xml.AddFirst(pPr = new XElement(Name.ParagraphProperties));
            return new ParagraphProperties(pPr);
        }
            
        private set
        {
            var pPr = Xml.Element(Name.ParagraphProperties);
            pPr?.Remove();

            var xml = value.Xml;
            if (xml.Parent != null)
                xml = xml.Clone();
            
            Xml.AddFirst(xml);
        }
    }

    /// <summary>
    /// Sets the properties on the paragraph with a fluent method.
    /// </summary>
    /// <param name="properties"></param>
    /// <returns></returns>
    public Paragraph WithProperties(ParagraphProperties properties)
    {
        Properties = properties;
        return this;
    }

    /// <summary>
    /// Returns a list of all Pictures in a paragraph.
    /// </summary>
    public IReadOnlyList<Picture> Pictures => (
        from p in Xml.LocalNameDescendants("pic")
        let id = p.FirstLocalNameDescendant("blip").AttributeValue(Namespace.RelatedDoc + "embed")
        where id != null
        select new Picture(SafeDocument, SafePackagePart, p,
            new Image(SafeDocument, SafeDocument?.SafePackagePart?.GetRelationship(id)))
    ).Union(
        from p in Xml.LocalNameDescendants("pict")
        let id = p.FirstLocalNameDescendant("imagedata").AttributeValue(Namespace.RelatedDoc + "id")
        where id != null
        select new Picture(SafeDocument, SafePackagePart, p,
            new Image(SafeDocument, SafeDocument?.SafePackagePart?.GetRelationship(id)))
    ).ToList().AsReadOnly();

    /// <summary>
    /// Gets the findText value of this paragraph.
    /// </summary>
    public string Text => HelperFunctions.GetText(Xml);

    /// <summary>
    /// Append findText to this paragraph.
    /// </summary>
    /// <param name="text">The findText to append.</param>
    /// <param name="formatting">Formatting for the findText run</param>
    /// <returns>This paragraph with the new findText appended.</returns>
    public Paragraph Append(string text, Formatting? formatting = null)
    {
        if (Text.Length == 0)
        {
            SetText(text, formatting);
        }
        else
        {
            Xml.Add(HelperFunctions.FormatInput(text, formatting?.Xml));
        }

        return this;
    }

    /// <summary>
    /// Appends a new bookmark to the paragraph
    /// </summary>
    /// <param name="bookmarkName">Bookmark name</param>
    /// <param name="firstRun">Start run to set bookmark on</param>
    /// <param name="lastRun">End run to set bookmark on</param>
    /// <returns>This paragraph</returns>
    public Paragraph SetBookmark(string bookmarkName, Run? firstRun = null, Run? lastRun = null)
    {
        if (this.BookmarkExists(bookmarkName))
            throw new ArgumentException($"Bookmark '{bookmarkName}' already exists.", nameof(bookmarkName));

        long id = InDocument
            ? Document.GetNextDocumentId()
            : 0;

        firstRun ??= Runs.FirstOrDefault();
        lastRun ??= Runs.LastOrDefault();

        var bookmarkStart = new XElement(Name.BookmarkStart, 
            new XAttribute(Name.Id, id), new XAttribute(Name.NameId, bookmarkName));

        var bookmarkEnd = new XElement(Name.BookmarkEnd,
            new XAttribute(Name.Id, id), new XAttribute(Name.NameId, bookmarkName));

        // Add the bookmarkStart before all runs.
        if (firstRun != null)
        {
            firstRun.Xml.AddBeforeSelf(bookmarkStart);
        }
        else
        {
            Xml.Add(bookmarkStart);
        }

        // .. and the bookmarkEnd after all runs.
        if (lastRun != null)
        {
            lastRun.Xml.AddAfterSelf(bookmarkEnd);            
        }
        else
        {
            Xml.Add(bookmarkEnd);
        }

        return this;
    }

    /// <summary>
    /// Add an equation to a document.
    /// </summary>
    /// <param name="equation">The Equation to append.</param>
    /// <returns>The paragraph with the Equation now appended.</returns>
    public Paragraph AppendEquation(string equation)
    {
        // Create equation element
        var oMathPara = new XElement(Name.MathParagraph,
            new XElement(Name.OfficeMath,
                new XElement(Namespace.Math + "r",
                    new Formatting { Font = new FontFamily("Cambria Math") }.Xml,
                    new XElement(Namespace.Math + "t", equation)
                )
            )
        );

        Xml.Add(oMathPara);

        return this;
    }

    /// <summary>
    /// This function inserts a hyperlink into a paragraph at a specified character index.
    /// </summary>
    /// <param name="hyperlink">The hyperlink to insert.</param>
    /// <param name="charIndex">The character index in the owning paragraph to insert at.</param>
    /// <returns>The paragraph with the Hyperlink inserted at the specified index.</returns>
    public Paragraph InsertHyperlink(Hyperlink hyperlink, int charIndex = 0)
    {
        if (InDocument)
        {
            hyperlink.SetOwner(Document, PackagePart, true);
            _ = hyperlink.GetOrCreateRelationship();
        }
        else
        {
            unownedHyperlinks.Add(hyperlink);
        }

        if (charIndex == 0)
        {
            // Add this hyperlink as the last element.
            Xml.AddFirst(hyperlink.Xml);
        }
        else
        {
            // Get the first run effected by this Insert
            var run = FindRunAffectedByEdit(EditType.Insert, charIndex);
            if (run == null)
            {
                // Add this hyperlink as the last element.
                Xml.Add(hyperlink.Xml);
            }
            else
            {
                // Split this run at the point you want to insert
                var splitRun = run.SplitAtIndex(charIndex);

                // Replace the original run.
                run.Xml.ReplaceWith(splitRun[0], hyperlink.Xml, splitRun[1]);
            }
        }

        return this;
    }

    /// <summary>
    /// Returns a list of Hyperlinks in this paragraph.
    /// </summary>
    public List<Hyperlink> Hyperlinks => Hyperlink.Enumerate(this, unownedHyperlinks).ToList();

    /// <summary>
    /// Append a hyperlink to a paragraph.
    /// </summary>
    /// <param name="hyperlink">The hyperlink to append.</param>
    /// <returns>The paragraph with the hyperlink appended.</returns>
    public Paragraph Append(Hyperlink hyperlink)
    {
        if (InDocument)
        {
            hyperlink.SetOwner(Document, PackagePart, true);
            _ = hyperlink.GetOrCreateRelationship();
        }
        else
        {
            unownedHyperlinks.Add(hyperlink);
        }

        Xml.Add(hyperlink.Xml);
        return this;
    }

    /// <summary>
    /// Append a PageCount place holder onto the end of a paragraph.
    /// </summary>
    /// <param name="format">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    public void AddPageCount(PageNumberFormat format) => AddPageNumberInfo(format, "numPages");

    /// <summary>
    /// Append a PageNumber place holder onto the end of a paragraph.
    /// </summary>
    /// <param name="format">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    public void AddPageNumber(PageNumberFormat format) => AddPageNumberInfo(format, "page");

    /// <summary>
    /// Insert a PageCount place holder into a paragraph.
    /// This place holder should only be inserted into a Header or Footer paragraph.
    /// Word will not automatically update this field if it is inserted into a document level paragraph.
    /// </summary>
    /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    /// <param name="index">The findText index to insert this PageCount place holder at.</param>
    public void InsertPageCount(PageNumberFormat pnf, int index = 0) => AddPageNumberInfo(pnf, "numPages", index);

    /// <summary>
    /// Insert a PageNumber place holder into a paragraph.
    /// This place holder should only be inserted into a Header or Footer paragraph.
    /// Word will not automatically update this field if it is inserted into a document level paragraph.
    /// </summary>
    /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    /// <param name="index">The findText index to insert this PageNumber place holder at.</param>
    public void InsertPageNumber(PageNumberFormat pnf, int index = 0) => AddPageNumberInfo(pnf, "page", index);

    /// <summary>
    /// Internal method to populate page numbers or page counts.
    /// </summary>
    /// <param name="format">Page number format</param>
    /// <param name="type">Numbers or Counts</param>
    /// <param name="index">Position to insert</param>
    private void AddPageNumberInfo(PageNumberFormat format, string type, int? index = null)
    {
        var fldSimple = new XElement(Name.SimpleField);

        fldSimple.Add(format == PageNumberFormat.Normal
            ? new XAttribute(Name.Instr, $@" {type.ToUpper()}   \* MERGEFORMAT ")
            : new XAttribute(Name.Instr, $@" {type.ToUpper()}  \* ROMAN  \* MERGEFORMAT "));

        var content = XElement.Parse(
            @"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof />
                   </w:rPr>
                   <w:t>1</w:t>
               </w:r>"
        );

        if (index is null or < 0)
        {
            fldSimple.Add(content);
            Xml.Add(fldSimple);
        }
        else if (index == 0)
        {
            Xml.AddFirst(fldSimple);
        }
        else
        {
            var r = FindRunAffectedByEdit(EditType.Insert, index.Value);
            if (r != null)
            {
                var splitEdit = SplitEdit(r.Xml, index.Value, EditType.Insert);
                r.Xml.ReplaceWith(splitEdit[0], fldSimple, splitEdit[1]);
            }
        }
    }

    /// <summary>
    /// Add an image to a document, create a custom view of that image (picture) and then insert it into a paragraph using append.
    /// </summary>
    /// <param name="drawing">The Picture to append.</param>
    /// <returns>The paragraph with the Picture now appended.</returns>
    public Paragraph Append(Drawing drawing)
    {
        if (InDocument)
        {
            drawing.SetOwner(Document, PackagePart, true);
            var picture = drawing.Picture ?? throw new InvalidOperationException("Failed to create picture from drawing - possibly missing image?"); ;
            picture.RelationshipId = picture.GetOrCreateImageRelationship();
        }

        // Add a new run with the given drawing to the paragraph.
        var run = new XElement(Name.Run,
            new XElement(Name.RunProperties,
                new XElement(Namespace.Main + "noProof")),
            drawing.Xml);
        Xml.Add(run);

        return this;
    }

    /// <summary>
    /// Find all instances of a string in this paragraph and return their indexes in a List.
    /// </summary>
    /// <param name="findText">The string to find</param>
    /// <param name="comparisonType">True to ignore case in the search</param>
    /// <returns>A list of indexes.</returns>
    public IEnumerable<int> FindAll(string findText, StringComparison comparisonType)
    {
        if (findText == null) throw new ArgumentNullException(nameof(findText));

        List<int> foundIndexes = new List<int>();
        var text = Text;
        if (!string.IsNullOrEmpty(text))
        {
            int start = text.IndexOf(findText, comparisonType);
            while (start >= 0)
            {
                foundIndexes.Add(start);
                start = text.IndexOf(findText, start + findText.Length, comparisonType);
            }
        }

        return foundIndexes;
    }

    /// <summary>
    ///  Find all unique instances of the given Regex Pattern
    /// </summary>
    /// <param name="regex">Regex to match</param>
    /// <returns>Index and matched findText</returns>
    public IEnumerable<(int index, string text)> FindPattern(Regex regex)
    {
        MatchCollection mc = regex.Matches(Text);
        return mc.Select(m => (index: m.Index, text: m.Value));
    }

    /// <summary>
    /// Retrieve all the bookmarks in this paragraph
    /// </summary>
    /// <returns>Enumerable of bookmark objects</returns>
    internal IEnumerable<Bookmark> GetBookmarks()
    {
        return Xml.Descendants(Name.BookmarkStart)
            .Select(e => new Bookmark(e, this));
    }

    /// <summary>
    /// Insert a findText block at a bookmark
    /// </summary>
    /// <param name="bookmarkName">Bookmark name</param>
    /// <param name="toInsert">Text to insert</param>
    public bool InsertAtBookmark(string bookmarkName, string toInsert)
    {
        var bookmark = GetBookmarks().SingleOrDefault(bm => bm.Name == bookmarkName);
        if (bookmark == null) 
            return false;
            
        var run = HelperFunctions.FormatInput(toInsert, null);
        bookmark.Xml.AddBeforeSelf(run);
        return true;
    }

    /// <summary>
    /// Insert a field of type document property, this field will display the property in the paragraph.
    /// </summary>
    /// <param name="name">Property name</param>
    /// <param name="formatting">The formatting to use for this findText.</param>
    public Paragraph AddDocumentPropertyField(DocumentPropertyName name, Formatting? formatting = null)
    {
        if (Document == null)
            throw new InvalidOperationException("Cannot add document properties without a document owner.");

        Document.DocumentProperties.TryGetValue(name, out var propertyValue);
        if (!string.IsNullOrEmpty(propertyValue))
        {
            _ = AddComplexField(name.ToString().ToUpperInvariant(), propertyValue, formatting);
        }

        return this;
    }

    /// <summary>
    /// Insert a field of type document property, this field will display the property in the paragraph.
    /// </summary>
    /// <param name="name">Property name</param>
    /// <param name="formatting">The formatting to use for this findText.</param>
    public Paragraph AddCustomPropertyField(string name, Formatting? formatting = null)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
        if (Document == null)
            throw new InvalidOperationException("Cannot add document properties without a document owner.");

        Document.CustomProperties.TryGetValue(name, out var propertyValue);
        _ = AddComplexField(name, propertyValue?.ToString(), formatting);

        return this;
    }

    /// <summary>
    /// Inserts a complex field into the paragraph
    /// </summary>
    /// <param name="name">Name of field</param>
    /// <param name="fieldValue">Value of field</param>
    /// <param name="formatting">Formatting to apply</param>
    /// <returns>Inserted DocProperty</returns>
    private DocProperty AddComplexField(string name, string? fieldValue, Formatting? formatting = null)
    {
        if (Document == null || PackagePart == null)
            throw new InvalidOperationException("Paragraph not part of document.");

        // Start of complex field
        var start = new XElement(Name.Run,
            new XElement(Name.RunProperties, new XElement(Namespace.Main + "noProof")),
            new XElement(Name.ComplexField, new XAttribute(Namespace.Main + "fldCharType", "begin"))
        );

        // Property definition
        var pdef = new XElement(Name.Run,
            new XElement(Name.RunProperties, new XElement(Namespace.Main + "noProof")),
            new XElement(Namespace.Main + "instrText", new XAttribute(XNamespace.Xml + "space", "preserve"),
                $@"{name} \* MERGEFORMAT ")
        );

        // Separator
        var sep = new XElement(Name.Run,
            new XElement(Name.RunProperties, new XElement(Namespace.Main + "noProof")),
            new XElement(Name.ComplexField, new XAttribute(Namespace.Main + "fldCharType", "separate"))
        );

        // Value
        formatting ??= new Formatting();
        formatting.NoProof = true;

        var value = new XElement(Name.Run,
            formatting.Xml,
            new XElement(Name.Text,
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                fieldValue ?? "")
        );

        // End marker
        var end = new XElement(Name.Run,
            new XElement(Name.RunProperties, new XElement(Namespace.Main + "noProof")),
            new XElement(Name.ComplexField, new XAttribute(Namespace.Main + "fldCharType", "end"))
        );

        Xml.Add(start, pdef, sep, value, end);
        return new DocProperty(Document, PackagePart, pdef, value);
    }

    //TODO: add simple field property
    /*
    /// <summary>
    /// Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
    /// </summary>
    /// <param name="name"></param>
    /// <param name="formatting">The formatting to use for this findText.</param>
    public DocProperty AddDocumentProperty2(DocumentPropertyName name, Formatting formatting = null)
    {
        if (Document == null)
            throw new InvalidOperationException("Cannot add document properties without a document owner.");

        var p = Document.DocumentProperties.SingleOrDefault(p => p.Name == name);
        XElement xml = new XElement(Name.SimpleField,
            new XAttribute(Name.Instr, $@"DOCPROPERTY {name.GetEnumName()} \* MERGEFORMAT"),
            new XElement(Name.Run, new XElement(Name.Text, formatting?.Xml, p.Value))
        );

        Xml.Add(xml);

        return new DocProperty(Document, xml);
    }
    */

    /// <summary>
    /// Insert a Picture into a paragraph at the given findText index.
    /// If not index is provided defaults to 0.
    /// </summary>
    /// <param name="picture">The Picture to insert.</param>
    /// <param name="index">The findText index to insert at.</param>
    /// <returns>The modified paragraph.</returns>
    public Paragraph Insert(Picture picture, int index = 0)
    {
        if (InDocument)
        {
            picture.SetOwner(Document, PackagePart, true);
            picture.RelationshipId = picture.GetOrCreateImageRelationship();
        }

        // Create a run for the picture
        var xml = new XElement(Name.Run,
            new XElement(Name.RunProperties,
                new XElement(Namespace.Main + "noProof")),
            picture.Xml);

        if (index == 0)
        {
            // Add this hyperlink as the last element.
            Xml.AddFirst(xml);
        }
        else
        {
            // Get the first run effected by this Insert
            var run = FindRunAffectedByEdit(EditType.Insert, index);
            if (run == null)
            {
                // Add this picture as the last element.
                Xml.Add(xml);
            }
            else
            {
                // Split this run at the point you want to insert
                var splitRun = run.SplitAtIndex(index);

                // Replace the original run.
                run.Xml.ReplaceWith(splitRun[0], xml, splitRun[1]);
            }
        }

        return this;
    }

    /// <summary>
    /// Inserts a string into a paragraph with the specified formatting.
    /// </summary>
    public void InsertText(string value, Formatting? formatting = null)
    {
        Xml.Add(HelperFunctions.FormatInput(value, formatting?.Xml));
    }

    /// <summary>
    /// Replaces any existing findText runs in this paragraph with the specified findText.
    /// </summary>
    /// <param name="value">Text value</param>
    /// <param name="formatting">Formatting to apply (null for default)</param>
    public void SetText(string value, Formatting? formatting = null)
    {
        // Remove all runs from this paragraph.
        Xml.Descendants(Name.Run).Remove();

        // Add the new run.
        if (!string.IsNullOrEmpty(value))
        {
            Xml.Add(HelperFunctions.FormatInput(value, formatting?.Xml));
        }
    }

    /// <summary>
    /// Inserts a string into the paragraph with the specified formatting at the given index.
    /// </summary>
    /// <param name="index">The index position of the insertion.</param>
    /// <param name="value">The System.String to insert.</param>
    /// <param name="formatting">The findText formatting.</param>
    public void InsertText(int index, string value, Formatting? formatting = null)
    {
        // Get the first run effected by this Insert
        var run = FindRunAffectedByEdit(EditType.Insert, index);
        if (run == null)
        {
            Xml.Add(HelperFunctions.FormatInput(value, formatting?.Xml));
        }
        else
        {
            object insert = HelperFunctions.FormatInput(value, formatting?.Xml);
            var parentElement = run.Xml.Parent;
            if (parentElement == null)
            {
                throw new InvalidOperationException("Orphaned run not connected to paragraph.");
            }

            switch (parentElement.Name.LocalName)
            {
                case "ins":
                case "del":
                    var splitEdit = SplitEdit(parentElement, index, EditType.Insert);
                    parentElement.ReplaceWith(splitEdit[0], insert, splitEdit[1]);
                    break;

                default:
                    var splitRun = run.SplitAtIndex(index);
                    run.Xml.ReplaceWith(splitRun[0], insert, splitRun[1]);
                    break;
            }
        }
    }

    /// <summary>
    /// Remove this paragraph from the document.
    /// </summary>
    public void Remove()
    {
        // If this is the only paragraph in the Cell then we cannot remove it.
        if (Xml.Parent?.Name.LocalName == "tc"
            && Xml.Parent.Elements(Name.Paragraph).Count() == 1)
        {
            Xml.Value = string.Empty;
        }
        else
        {
            Xml.Remove();
        }
    }

    /// <summary>
    /// Removes characters from a DXPlus.Document.paragraph.
    /// </summary>
    /// <param name="index">The position to begin deleting characters.</param>
    /// <param name="count">The number of characters to delete</param>
    public void RemoveText(int index, int count)
    {
        // The number of characters processed so far
        int processed = 0;

        do
        {
            // Get the first run effected by this Remove
            var run = FindRunAffectedByEdit(EditType.Delete, index);

            // The parent of this Run
            var parentElement = run?.Xml.Parent;
            switch (parentElement?.Name.LocalName)
            {
                case "ins":
                case "del":
                {
                    var splitEditBefore = SplitEdit(parentElement, index, EditType.Delete);
                    int take = count - processed;
                    var splitEditAfter = SplitEdit(parentElement, index + take, EditType.Delete);
                    var before = splitEditBefore[1];
                    Debug.Assert(before != null);
                    var middle = SplitEdit(before, index + take, EditType.Delete)[1];
                    processed += HelperFunctions.GetTextLength(middle);
                    parentElement.ReplaceWith(splitEditBefore[0], null, splitEditAfter[1]);
                }
                    break;

                default:
                    if (run != null && HelperFunctions.GetTextLength(run.Xml) > 0)
                    {
                        var splitRunBefore = run.SplitAtIndex(index);
                        int min = Math.Min(index + (count - processed), run.EndIndex);
                        var splitRunAfter = run.SplitAtIndex(min);
                        var middle = new Run(SafeDocument, SafePackagePart, splitRunBefore[1]!, run.StartIndex + HelperFunctions.GetTextLength(splitRunBefore[0])).SplitAtIndex(min)[0];
                        processed += HelperFunctions.GetTextLength(middle);
                        run.Xml.ReplaceWith(splitRunBefore[0], null, splitRunAfter[1]);
                    }
                    else
                    {
                        processed = count;
                    }

                    break;
            }

            // See if the paragraph is empty -- if so we can remove it.
            if (parentElement != null
                && HelperFunctions.GetTextLength(parentElement) == 0
                && parentElement.Parent != null
                && parentElement.Parent.Name.LocalName != "tc"
                && parentElement.Parent.Elements(Name.Paragraph).Any()
                && !parentElement.Descendants(Namespace.Main + "drawing").Any())
            {
                parentElement.Remove();
            }
        }
        while (processed < count);
    }

    /// <summary>
    /// Replaces all occurrences of a specified findText in this instance.
    /// </summary>
    /// <param name="findText">Regular expression to search for</param>
    /// <param name="replaceText">Text to replace all occurrences of oldValue.</param>
    /// <param name="comparisonType">Comparison type - defaults to current culture</param>
    public bool FindReplace(string findText, string? replaceText, StringComparison comparisonType = StringComparison.CurrentCulture)
    {
        if (findText == null) throw new ArgumentNullException(nameof(findText));

        if (string.IsNullOrEmpty(Text)) 
            return false;

        int start = Text.IndexOf(findText, comparisonType);
        bool found = start >= 0;
        while (start >= 0)
        {
            RemoveText(start, findText.Length);
            if (!string.IsNullOrEmpty(replaceText))
                InsertText(start, replaceText);

            start = Text.IndexOf(findText, start+replaceText?.Length??0, comparisonType);
        }

        return found;
    }

    /// <summary>
    /// Walk all the findText runs in the paragraph and find the one containing a specific index.
    /// </summary>
    /// <param name="editType">Type of edit being performed (insert or delete)</param>
    /// <param name="index">Index to look for</param>
    /// <returns>Run containing index</returns>
    internal Run? FindRunAffectedByEdit(EditType editType, int index)
    {
        int len = HelperFunctions.GetText(Xml).Length;
        if (index < 0 || (editType == EditType.Insert && index > len) || (editType == EditType.Delete && index >= len))
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        int count = 0; 
        Run? run = null;
        
        RecursiveSearchForRunByIndex(Xml, editType, index, ref count, ref run);

        return run;
    }

    /// <summary>
    /// Recursive method to identify a findText run from a starting element and index.
    /// </summary>
    /// <param name="el">Element to search</param>
    /// <param name="editType">Type of edit being performed (insert or delete)</param>
    /// <param name="index">Index to look for</param>
    /// <param name="count">Total searched</param>
    /// <param name="run">The located findText run</param>
    private void RecursiveSearchForRunByIndex(XElement el, EditType editType, int index, ref int count, ref Run? run)
    {
        count += HelperFunctions.GetSize(el);
        if (count > 0 && (editType == EditType.Delete && count > index || editType == EditType.Insert && count >= index))
        {
            // Correct the index
            count -= el.ElementsBeforeSelf().Sum(HelperFunctions.GetSize);
            count -= HelperFunctions.GetSize(el);
            count = Math.Max(0, count);

            // We have found the element, now find the run it belongs to.
            var search = el.FindParent(Name.Run);
            if (search == null) return;
            run = new Run(SafeDocument, SafePackagePart, search, count);
        }
        else if (el.HasElements)
        {
            foreach (var e in el.Elements())
            {
                if (run == null)
                {
                    // Run can be changed by this method.
                    RecursiveSearchForRunByIndex(e, editType, index, ref count, ref run);
                }
            }
        }
    }

    /// <summary>
    /// Splits a tracked edit (ins/del)
    /// </summary>
    /// <param name="element">Parent element to split</param>
    /// <param name="index">Character index to split on</param>
    /// <param name="type">Type of edit being performed (insert/delete)</param>
    /// <returns>Split XElement array</returns>
    internal XElement?[] SplitEdit(XElement element, int index, EditType type)
    {
        // Find the run containing the index
        var run = FindRunAffectedByEdit(type, index);
        if (run == null)
            return new XElement?[] {null, null};

        var splitRun = run.SplitAtIndex(index);

        XElement? splitLeft = new(element.Name, element.Attributes(), run.Xml.ElementsBeforeSelf(), splitRun[0]);
        if (HelperFunctions.GetTextLength(splitLeft) == 0)
        {
            splitLeft = null;
        }

        XElement? splitRight = new(element.Name, element.Attributes(), splitRun[1], run.Xml.ElementsAfterSelf());
        if (HelperFunctions.GetTextLength(splitRight) == 0)
        {
            splitRight = null;
        }

        return new[] { splitLeft, splitRight };
    }

    /// <summary>
    /// Called when the document owner is changed.
    /// </summary>
    protected override void OnAddToDocument()
    {
        // Update bookmark IDs.
        foreach (var bookmark in GetBookmarks())
            bookmark.Id = Document.GetNextDocumentId();

        if (Hyperlinks.Any())
        {
            // Ensure the owner document has the hyperlink styles.
            Document.AddHyperlinkStyle();
            _ = Document.EnsureRelsPathExists(PackagePart);

            // Fix up hyperlinks
            foreach (var hyperlink in Hyperlinks)
            {
                hyperlink.SetOwner(Document, PackagePart,true);
                _ = hyperlink.GetOrCreateRelationship();
            }

            unownedHyperlinks.Clear();
        }

        if (tableAfterParagraph != null)
        {
            tableAfterParagraph.SetOwner(Document, PackagePart, true);
            Xml.AddAfterSelf(tableAfterParagraph.Xml);
            tableAfterParagraph = null;
        }

        if (Pictures.Any())
        {
            // Check to see if the .rels file exists and create it if not.
            _ = Document.EnsureRelsPathExists(PackagePart);

            // Fix up pictures
            foreach (var picture in Pictures)
            {
                picture.SetOwner(Document, PackagePart, true);
                picture.RelationshipId = picture.GetOrCreateImageRelationship();
            }
        }
    }

    /// <summary>
    /// This is used to change the start/end index for this paragraph object
    /// when it's inserted into a container.
    /// </summary>
    /// <param name="index">New starting index</param>
    internal void SetStartIndex(int index)
    {
        StartIndex = index;
        EndIndex = index + HelperFunctions.GetTextLength(Xml);
    }

    /// <summary>
    /// Create a new paragraph from some findText
    /// </summary>
    /// <param name="text"></param>
    /// <param name="formatting"></param>
    /// <returns>New paragraph</returns>
    internal static XElement Create(string text, Formatting? formatting)
    {
        return new XElement(Name.Paragraph,
            new XAttribute(Name.ParagraphId, HelperFunctions.GenerateHexId()),
            HelperFunctions.FormatInput(text, formatting?.Xml));
    }

    /// <summary>
    /// Method to clone a paragraph into a new unowned paragraph.
    /// </summary>
    /// <param name="otherParagraph">Existing paragraph</param>
    /// <returns></returns>
    public static Paragraph Clone(Paragraph otherParagraph)
    {
        if (otherParagraph is null)
            throw new ArgumentNullException(nameof(otherParagraph));

        return new Paragraph {
            Xml = otherParagraph.Xml.Clone(),
            Id = HelperFunctions.GenerateHexId()
        };
    }

    /// <summary>
    /// Determines equality for paragraphs
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Paragraph? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);
}