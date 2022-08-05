using System.Diagnostics;
using System.Drawing;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DXPlus.Charts;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a document paragraph.
/// </summary>
[DebuggerDisplay("{" + nameof(Text) + "}")]
public sealed class Paragraph : Block, IEquatable<Paragraph>
{
    private Table? tableAfterParagraph;
    private readonly List<Hyperlink> unownedHyperlinks = new();
    private readonly List<Chart> unownedCharts = new();

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
            // Only look at the localName so we capture Math.r and Main.r
            return Xml.Descendants()
                .Where(x => x.Name.LocalName == Name.Run.LocalName)
                .Select(runXml => new Run(SafeDocument, SafePackagePart, runXml));
        }
    }

    /// <summary>
    /// Create a paragraph from a string.
    /// </summary>
    /// <param name="text"></param>
    public static implicit operator Paragraph(string text) => new(text);

    /// <summary>
    /// Create a paragraph from a Run.
    /// </summary>
    /// <param name="text"></param>
    public static implicit operator Paragraph(Run text) => new(text);

    /// <summary>
    /// Public constructor for the paragraph
    /// </summary>
    public Paragraph() 
        : this(null, null, new XElement(Name.Paragraph))
    {
    }

    /// <summary>
    /// Public constructor for the paragraph
    /// </summary>
    /// <param name="text"></param>
    public Paragraph(Run text) : this()
    {
        if (text == null) throw new ArgumentNullException(nameof(text));
        Xml.Add(text.Xml.Normalize());
    }

    /// <summary>
    /// Public constructor for the paragraph
    /// </summary>
    /// <param name="runs"></param>
    public Paragraph(IEnumerable<Run> runs) : this()
    {
        if (runs == null) throw new ArgumentNullException(nameof(runs));
        AddRange(runs);
    }

    /// <summary>
    /// Public constructor for the paragraph
    /// </summary>
    /// <param name="text">Text to add</param>
    public Paragraph(string text) : this(new Run(text))
    {
    }

    /// <summary>
    /// Public constructor for the paragraph
    /// </summary>
    /// <param name="text">Text to add</param>
    /// <param name="formatting">Formatting to apply</param>
    public Paragraph(string text, Formatting formatting) : this(new Run(text, formatting))
    {
    }

    /// <summary>
    /// Constructor for the paragraph
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package owner</param>
    /// <param name="xml">XML for the paragraph</param>
    internal Paragraph(Document? document, PackagePart? packagePart, XElement xml) : base(xml)
    {
        if (document != null)
            SetOwner(document, packagePart, false);
    }

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
    /// Clears all formatting tied to this paragraph.
    /// </summary>
    /// <returns>This paragraph</returns>
    public Paragraph ClearFormatting()
    {
        Runs.ToList().ForEach(r => r.Properties = null);
        DefaultFormatting = null;
        return this;
    }

    /// <summary>
    /// Returns a list of DocProperty elements in this document.
    /// </summary>
    public IEnumerable<DocProperty> Fields
    {
        get
        {
            if (!InDocument) return Enumerable.Empty<DocProperty>();

            var properties = Xml.Descendants(Name.SimpleField)
                .Select(el => new DocProperty(Document, Document.PackagePart, el, null)).ToList();

            // Look for complex field insertions in the paragraph. These should always be in run elements and
            // have a start, name, sep, value section.
            foreach (var field in Xml.Descendants(Name.ComplexField)
                         .Where(e => e.AttributeValue(Namespace.Main + "fldCharType") == "begin")
                         .ToList())
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

                // Look for the next [w:t] text element.
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
            /*
            var ps = Document.Paragraphs.ToList();
            int pos = ps.IndexOf(this);
            return pos > 0 ? ps[pos - 1] : null;
            */

            var previous = this.Xml.PreviousSibling(Name.Paragraph);
            if (previous == null) return null;

            int pos = 0;
            return DocumentHelpers.WrapParagraphElement(previous, this.Document, this.PackagePart, ref pos);
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

            /*
            var ps = Document.Paragraphs.ToList();
            int pos = ps.IndexOf(this);
            return pos >= 0 && pos < ps.Count-1 ? ps[pos+1] : null;
            */

            var next = this.Xml.NextSibling(Name.Paragraph);
            if (next == null) return null;

            int pos = 0;
            return DocumentHelpers.WrapParagraphElement(next, this.Document, this.PackagePart, ref pos);
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
    /// Optimize the runs in this paragraph by collapsing all adjacent w:r elements with the same formatting
    /// into a single run.
    /// </summary>
    public void OptimizeRuns()
    {
        Run? currentRun = null;

        var children = Xml.Elements().Where(e => e.Name != Name.ParagraphProperties).ToList();
        foreach (var child in children)
        {
            if (child.Name != Name.Run)
            {
                currentRun = null;
                continue;
            }

            if (currentRun == null)
                currentRun = new Run(null, null, child);
            else
            {
                // See if we can merge.
                var testRun = new Run(null, null, child);
                if (testRun.HasText 
                    && (testRun.Properties == null && currentRun.Properties == null
                    || testRun.Properties?.Equals(currentRun.Properties) == true))
                {
                    var insertText = testRun.Xml.Element(Name.Text)!;
                    var text = currentRun.Xml.Element(Name.Text);
                    if (text == null)
                    {
                        currentRun.Xml.Add(insertText);
                    }
                    else
                    {
                        text.Value += insertText.Value;
                        text.PreserveSpace();
                    }
                    child.Remove();
                }
                else
                {
                    currentRun = testRun;
                }
            }
        }
    }

    /// <summary>
    /// Create a new paragraph, append it to the document and add the specified chart to it
    /// </summary>
    public Paragraph Add(string name, Chart chart)
    {
        string relationId;
        long chartId;

        if (InDocument)
        {
            (relationId, chartId) = Document.ChartManager.CreateRelationship(chart);
        }
        else
        {
            unownedCharts.Add(chart);
            relationId = (unownedCharts.Count * -1).ToString();
            chartId = 0;
        }



        // Create the XML needed to host the chart in a run.
        var chartElement = new XElement(Name.Run,
            new XElement(Name.RunProperties, new XElement(Name.NoProof)),
            new XElement(Namespace.Main + RunTextType.Drawing,
                new XElement(Namespace.WordProcessingDrawing + "inline",
                    new XAttribute("distB", 0), 
                    new XAttribute("distL", 0),
                    new XAttribute("distR", 0),
                    new XAttribute("distT", 0),
                    new XElement(Namespace.WordProcessingDrawing + "extent",
                        new XAttribute("cx", 576 * Uom.EmuConversion),
                        new XAttribute("cy", 336 * Uom.EmuConversion)),
                    new XElement(Namespace.WordProcessingDrawing + "effectExtent",
                        new XAttribute("l", 0),
                        new XAttribute("t", 0),
                        new XAttribute("r", 0),
                        new XAttribute("b", 0)),
                    new XElement(Namespace.WordProcessingDrawing + "docPr",
                        new XAttribute("id", chartId),
                        new XAttribute("name", name)),
                    new XElement(Namespace.WordProcessingDrawing + "cNvGraphicFramePr"),
                    new XElement(Namespace.DrawingMain + "graphic",
                        new XElement(Namespace.DrawingMain + "graphicData",
                            new XAttribute("uri", Namespace.Chart.NamespaceName),
                            new XElement(Namespace.Chart + "chart",
                                new XAttribute(Namespace.RelatedDoc + "id", relationId)
                            )
                        )
                    )
                )
            ));
        
        // Add it to our paragraph
        Xml.Add(chartElement);
        return this;
    }


    /// <summary>
    /// Insert a table before this paragraph
    /// </summary>
    /// <param name="table"></param>
    public Paragraph InsertBefore(Table table)
    {
        if (table == null) throw new ArgumentNullException(nameof(table));
        if (table.Xml.HasParent())
            throw new ArgumentException("Cannot add table multiple times.", nameof(table));
        if (!InDocument)
            throw new InvalidOperationException("Cannot insert table without owning document.");

        table.SetOwner(Document, PackagePart, true);
        Xml.AddBeforeSelf(table.Xml);

        return this;
    }

    /// <summary>
    /// Add a table after this paragraph.
    /// </summary>
    /// <param name="table">Table to add</param>
    public Paragraph InsertAfter(Table table)
    {
        if (table == null) throw new ArgumentNullException(nameof(table));
        if (table.Xml.HasParent())
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
            
        set
        {
            var pPr = Xml.Element(Name.ParagraphProperties);
            pPr?.Remove();

            var xml = value.Xml!;
            if (xml.Parent != null)
                xml = xml.Clone();
            
            Xml.AddFirst(xml);
        }
    }

    /// <summary>
    /// Returns a list of all Pictures in a paragraph.
    /// </summary>
    public IEnumerable<Drawing> Drawings 
        => Xml.Descendants(Name.Drawing).Select(e => new Drawing(SafeDocument, SafePackagePart, e));

    /// <summary>
    /// Gets or replaces the text value of this paragraph.
    /// </summary>
    public string Text
    {
        get => DocumentHelpers.GetText(Xml, false);
        set
        {
            Xml.Descendants(Name.Run).Remove();
            if (!string.IsNullOrEmpty(value))
                Xml.Add(DocumentHelpers.CreateRunElements(value, null));
        }
    }

    /// <summary>
    /// Add a run of text to this paragraph
    /// </summary>
    /// <param name="run">Run of text to add</param>
    /// <returns>Paragraph with the text added</returns>
    public Paragraph AddText(Run run)
    {
        var child = run.InDocument ? run.Xml.Clone() : run.Xml;
        Xml.Add(child);
        return this;
    }

    /// <summary>
    /// Add a run of text to this paragraph
    /// </summary>
    /// <param name="text">Run of text to add</param>
    /// <param name="formatting">Formatting to merge into the run</param>
    /// <returns>Paragraph with the text added</returns>
    public Paragraph AddText(string text, Formatting formatting)
    {
        DocumentHelpers.CreateRunElements(text, formatting.Xml)
            .ToList()
            .ForEach(Xml.Add);
        
        return this;
    }

    /// <summary>
    /// Add a range of runs to the paragraph
    /// </summary>
    /// <param name="runs">Runs to add</param>
    /// <returns>Paragraph</returns>
    public Paragraph AddRange(IEnumerable<Run> runs)
    {
        if (runs == null) throw new ArgumentNullException(nameof(runs));
        foreach (var run in runs)
            AddText(run);
        return this;
    }

    /// <summary>
    /// Fluent method to add a bookmark to the end of the paragraph
    /// </summary>
    /// <param name="bookmarkName">Name</param>
    /// <returns>Paragraph owner</returns>
    public Paragraph AddBookmark(string bookmarkName)
    {
        if (this.BookmarkExists(bookmarkName))
            throw new ArgumentException($"Bookmark '{bookmarkName}' already exists.", nameof(bookmarkName));

        long id = InDocument
            ? Document.GetNextDocumentId()
            : 0;

        var bookmarkStart = new XElement(Name.BookmarkStart,
            new XAttribute(Name.Id, id), new XAttribute(Name.NameId, bookmarkName));

        var bookmarkEnd = new XElement(Name.BookmarkEnd,
            new XAttribute(Name.Id, id), new XAttribute(Name.NameId, bookmarkName));

        Xml.Add(bookmarkStart, bookmarkEnd);
        
        return this;
    }

    /// <summary>
    /// Appends a new bookmark to the paragraph
    /// </summary>
    /// <param name="bookmarkName">Bookmark name</param>
    /// <param name="firstRun">Start run to set bookmark on</param>
    /// <param name="lastRun">End run to set bookmark on</param>
    public void SetBookmark(string bookmarkName, Run? firstRun = null, Run? lastRun = null)
    {
        if (this.BookmarkExists(bookmarkName))
            throw new ArgumentException($"Bookmark '{bookmarkName}' already exists.", nameof(bookmarkName));

        long id = InDocument
            ? Document.GetNextDocumentId()
            : 0;

        firstRun ??= Runs.FirstOrDefault();
        lastRun ??= Runs.LastOrDefault();

        // Make sure we never insert bookmark start/end out of order.
        if (firstRun == null && lastRun != null)
            firstRun = lastRun;

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
    }

    /// <summary>
    /// Add an equation to a document.
    /// </summary>
    /// <param name="equation">The Equation to append.</param>
    /// <returns>The paragraph with the Equation now appended.</returns>
    public Paragraph AddEquation(string equation) => AddEquation(equation, new Formatting());

    /// <summary>
    /// Add an equation to a document.
    /// </summary>
    /// <param name="equation">The Equation to append.</param>
    /// <param name="formatting">Additional formatting to apply</param>
    /// <returns>The paragraph with the Equation now appended.</returns>
    public Paragraph AddEquation(string equation, Formatting formatting)
    {
        formatting.Font = new FontValue("Cambria Math");

        // Create equation element
        var oMathPara = new XElement(Name.MathParagraph,
            new XElement(Name.OfficeMath,
                new XElement(Namespace.Math + "r",
                    formatting.Xml,
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
    /// <param name="index">The character index in the owning paragraph to insert at.</param>
    /// <param name="hyperlink">The hyperlink to insert.</param>
    /// <returns>The paragraph with the Hyperlink inserted at the specified index.</returns>
    public Paragraph Insert(int index, Hyperlink hyperlink)
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

        if (index == 0)
        {
            // Add this hyperlink as the first element.
            Xml.AddFirst(hyperlink.Xml);
        }
        else
        {
            if (index >= DocumentHelpers.GetTextLength(Xml))
            {
                // Add this hyperlink as the last element.
                Xml.Add(hyperlink.Xml);
            }

            // Get the first run effected by this Insert
            var (run, startIndex) = FindRunAffectedByEdit(index);

            // Split this run at the point you want to insert
            var (leftElement, rightElement) = run.Split(index - startIndex);

            // Replace the original run.
            run.Xml.ReplaceWith(leftElement, hyperlink.Xml, rightElement);
        }

        return this;
    }

    /// <summary>
    /// Returns a list of Hyperlinks in this paragraph.
    /// </summary>
    public IEnumerable<Hyperlink> Hyperlinks => Hyperlink.Enumerate(this, unownedHyperlinks);

    /// <summary>
    /// Append a hyperlink to a paragraph.
    /// </summary>
    /// <param name="hyperlink">The hyperlink to append.</param>
    /// <returns>The paragraph with the hyperlink appended.</returns>
    public Paragraph Add(Hyperlink hyperlink)
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
    public void AddPageCount(PageNumberFormat format) => InsertPageNumberInfo(format, "numPages");

    /// <summary>
    /// Insert a PageCount place holder into a paragraph.
    /// This place holder should only be inserted into a Header or Footer paragraph.
    /// Word will not automatically update this field if it is inserted into a document level paragraph.
    /// </summary>
    /// <param name="index">The text index to insert this PageCount place holder at.</param>
    /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    public void InsertPageCount(int index, PageNumberFormat pnf) => InsertPageNumberInfo(pnf, "numPages", index);

    /// <summary>
    /// Append a PageNumber place holder onto the end of a paragraph.
    /// </summary>
    /// <param name="format">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    public void AddPageNumber(PageNumberFormat format) => InsertPageNumberInfo(format, "page");

    /// <summary>
    /// Insert a PageNumber place holder into a paragraph.
    /// This place holder should only be inserted into a Header or Footer paragraph.
    /// Word will not automatically update this field if it is inserted into a document level paragraph.
    /// </summary>
    /// <param name="index">The text index to insert this PageNumber place holder at.</param>
    /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
    public void InsertPageNumber(int index, PageNumberFormat pnf) => InsertPageNumberInfo(pnf, "page", index);

    /// <summary>
    /// Internal method to populate page numbers or page counts.
    /// </summary>
    /// <param name="format">Page number format</param>
    /// <param name="type">Numbers or Counts</param>
    /// <param name="index">Position to insert</param>
    private void InsertPageNumberInfo(PageNumberFormat format, string type, int? index = null)
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
            var (run, startIndex) = FindRunAffectedByEdit(index.Value);
            var (leftElement, rightElement) = run.Split(index.Value - startIndex);
            run.Xml.ReplaceWith(leftElement, fldSimple, rightElement);
        }
    }

    /// <summary>
    /// Add a picture into the document.
    /// </summary>
    /// <param name="picture">Picture to add</param>
    /// <returns>Paragraph owner</returns>
    public Paragraph Add(Picture picture) => Add(picture.Drawing!);

    /// <summary>
    /// Add a drawing into the document.
    /// </summary>
    /// <param name="drawing">The drawing to append.</param>
    /// <returns>The paragraph with the drawing now appended.</returns>
    public Paragraph Add(Drawing drawing)
    {
        if (InDocument)
        {
            drawing.SetOwner(Document, PackagePart, true);
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
    public IEnumerable<int> FindText(string findText, StringComparison comparisonType)
    {
        if (findText == null) throw new ArgumentNullException(nameof(findText));

        var foundIndexes = new List<int>();
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
    /// <returns>Index and matched text</returns>
    public IEnumerable<(int index, string text)> FindPattern(Regex regex)
    {
        if (regex == null) throw new ArgumentNullException(nameof(regex));
        return regex.Matches(Text).Select(m => (index: m.Index, text: m.Value));
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
    /// Insert a text block at a bookmark
    /// </summary>
    /// <param name="bookmarkName">Bookmark name</param>
    /// <param name="toInsert">Text to insert</param>
    public bool InsertTextAtBookmark(string bookmarkName, string toInsert)
    {
        var bookmark = GetBookmarks().SingleOrDefault(bm => bm.Name == bookmarkName);
        if (bookmark == null) 
            return false;
            
        var run = DocumentHelpers.CreateRunElements(toInsert, null);
        bookmark.Xml!.AddBeforeSelf(run);
        return true;
    }

    /// <summary>
    /// Insert a field of type document property, this field will display the property in the paragraph.
    /// </summary>
    /// <param name="name">Property name</param>
    /// <param name="formatting">The formatting to use for this text.</param>
    public Paragraph AddDocumentPropertyField(DocumentPropertyName name, Formatting? formatting = null)
    {
        if (!InDocument)
            throw new InvalidOperationException("Cannot add fields without a document owner.");

        var propertyValue = Document.Properties.GetValue(name);
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
    /// <param name="formatting">The formatting to use for this text.</param>
    public Paragraph AddCustomPropertyField(string name, Formatting? formatting = null)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
        if (!InDocument)
            throw new InvalidOperationException("Cannot add fields without a document owner.");

        Document.CustomProperties.TryGetValue(name, out var propertyValue);
        _ = AddComplexField(name, propertyValue?.Value, formatting);

        return this;
    }

    /// <summary>
    /// Inserts a complex field into the paragraph
    /// </summary>
    /// <param name="name">Name of field</param>
    /// <param name="fieldValue">Value of field</param>
    /// <param name="formatting">Formatting to apply</param>
    /// <returns>Inserted DocProperty</returns>
    private DocProperty AddComplexField(string name, string? fieldValue, Formatting? formatting)
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

    /// <summary>
    /// Insert a picture into the paragraph
    /// </summary>
    /// <param name="index">Zero-based index to insert at.</param>
    /// <param name="picture">Picture to insert</param>
    /// <returns>Paragraph owner</returns>
    public Paragraph Insert(int index, Picture picture) => Insert(index, picture.Drawing!);

    /// <summary>
    /// Insert a Drawing into a paragraph at the given text index.
    /// </summary>
    /// <param name="index">Zero-based index to insert at.</param>
    /// <param name="drawing">The drawing to insert.</param>
    /// <returns>The modified paragraph.</returns>
    public Paragraph Insert(int index, Drawing drawing)
    {
        if (InDocument)
        {
            drawing.SetOwner(Document, PackagePart, true);
        }

        // Create a run for the drawing
        var xml = new XElement(Name.Run,
            new XElement(Name.RunProperties,
                new XElement(Namespace.Main + "noProof")),
            drawing.Xml);

        if (index == 0)
        {
            // Add this hyperlink as the last element.
            Xml.AddFirst(xml);
        }
        else if (index >= DocumentHelpers.GetTextLength(Xml))
        {
            Xml.Add(xml);
        }
        else
        {
            // Get the first run effected by this Insert
            var (run, startIndex) = FindRunAffectedByEdit(index);
            // Split this run at the point you want to insert
            var (leftElement, rightElement) = run.Split(index - startIndex);

            // Replace the original run.
            run.Xml.ReplaceWith(leftElement, xml, rightElement);
        }

        return this;
    }

    /// <summary>
    /// Inserts a string into the paragraph at the specified index.
    /// </summary>
    /// <param name="index">The index position of the insertion.</param>
    /// <param name="run">The text to insert.</param>
    public Paragraph InsertText(int index, Run run)
    {
        InsertRuns(index, new [] { run.InDocument ? run.Xml.Clone() : run.Xml });
        return this;
    }

    /// <summary>
    /// Inserts a string into the paragraph at the specified index.
    /// </summary>
    /// <param name="index">The index position of the insertion.</param>
    /// <param name="text">The text to insert.</param>
    public Paragraph InsertText(int index, string text)
    {
        // Create the runs.
        var runsXml = DocumentHelpers.CreateRunElements(text, null);
        InsertRuns(index, runsXml);
        return this;
    }

    /// <summary>
    /// Inserts a string into the paragraph with the specified formatting at the given index.
    /// </summary>
    /// <param name="index">The index position of the insertion.</param>
    /// <param name="value">The System.String to insert.</param>
    /// <param name="formatting">The text formatting.</param>
    public Paragraph InsertText(int index, string value, Formatting formatting)
    {
        // Create the runs.
        var runsXml = DocumentHelpers.CreateRunElements(value, formatting.Xml);
        InsertRuns(index, runsXml);
        return this;
    }

    /// <summary>
    /// Insert a set of runs into a specified location in this paragraph.
    /// The existing run at the given index will be broken if necessary
    /// </summary>
    /// <param name="index">Index to insert at</param>
    /// <param name="elements">Elements to insert</param>
    /// <returns></returns>
    private void InsertRuns(int index, IEnumerable<XElement> elements)
    {
        if (index < 0 || index > DocumentHelpers.GetTextLength(Xml))
            throw new ArgumentOutOfRangeException(nameof(index));

        // Get the first run effected by this Insert
        var (run, startIndex) = FindRunAffectedByEdit(index);
        var parentElement = run.Xml.Parent!;

        if (parentElement.Name != Name.Paragraph)
        {
            var (leftElement, rightElement) = Split(parentElement, index);
            parentElement.ReplaceWith(leftElement, elements, rightElement);
        }
        else
        {
            var (leftElement, rightElement) = run.Split(index - startIndex);
            run.Xml.ReplaceWith(leftElement, elements, rightElement);
        }

        OptimizeRuns();
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
    /// Removes characters from this Paragraph.
    /// </summary>
    /// <param name="index">The position to begin deleting characters.</param>
    /// <param name="count">The number of characters to delete</param>
    public Paragraph RemoveText(int index, int count)
    {
        // The number of characters processed so far
        int processed = 0;

        do
        {
            // Get the run affected by this edit
            var (run, startIndex) = FindRunAffectedByEdit(index);

            // Get the parent element of the run
            var parentElement = run.Xml.Parent!;

            // Based on the parent of the run (direct paragraph vs. insert/delete marker), split
            // the run at the index point so we can begin removing the text.
            if (parentElement.Name != Name.Paragraph)
            {
                var (leftElement, rightElement) = Split(parentElement, index);
                int take = count - processed;
                var (_, after) = Split(parentElement, index + take);
                Debug.Assert(rightElement != null);
                var remove = Split(rightElement, index + take).rightElement;
                processed += DocumentHelpers.GetTextLength(remove);
                parentElement.ReplaceWith(leftElement, null, after);
            }
            else
            {
                var (leftElement, rightElement) = run.Split(index - startIndex);
                int endIndex = startIndex + DocumentHelpers.GetTextLength(run.Xml);
                int min = Math.Min(index + (count - processed), endIndex) - startIndex;
                var (_, splitRunAfter) = run.Split(min);
                var removeElement = new Run(SafeDocument, SafePackagePart, rightElement!).Split(count).leftElement ?? rightElement;
                processed += DocumentHelpers.GetTextLength(removeElement);
                run.Xml.ReplaceWith(leftElement, null, splitRunAfter);
            }

            // See if the paragraph is empty -- if so we can remove it.
            if (DocumentHelpers.GetTextLength(parentElement) == 0
                && parentElement.Parent != null
                && parentElement.Parent.Name.LocalName != "tc"
                && parentElement.Parent.Elements(Name.Paragraph).Any()
                && !parentElement.Descendants(Namespace.Main + RunTextType.Drawing).Any())
            {
                parentElement.Remove();
            }
        }
        while (processed < count);

        OptimizeRuns();

        return this;
    }

    /// <summary>
    /// Replaces all occurrences of a specified text in this instance.
    /// </summary>
    /// <param name="text">Regular expression to search for</param>
    /// <param name="replaceText">Text to replace all occurrences of oldValue.</param>
    /// <param name="comparisonType">Comparison type - defaults to current culture</param>
    public bool FindReplace(string text, string? replaceText, StringComparison comparisonType)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));

        if (string.IsNullOrEmpty(Text)) 
            return false;

        int start = Text.IndexOf(text, comparisonType);
        bool found = start >= 0;
        while (start >= 0)
        {
            RemoveText(start, text.Length);
            if (!string.IsNullOrEmpty(replaceText))
            {
                InsertText(start, replaceText);
            }

            start = Text.IndexOf(text, start+replaceText?.Length??0, comparisonType);
        }

        return found;
    }

    /// <summary>
    /// Split the paragraph at a specific character index
    /// </summary>
    /// <param name="index">Character index to split at</param>
    /// <returns>Left/Right split</returns>
    /// <exception cref="ArgumentNullException"></exception>
    internal (XElement? leftElement, XElement? rightElement) Split(int index)
    {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));

        var (run, startIndex) = FindRunAffectedByEdit(index);

        XElement? before, after;
        if (run.Xml.Parent?.Name.LocalName != Name.Paragraph.LocalName)
        {
            var (leftElement, rightElement) = Split(run.Xml.Parent!, index);
            before = new XElement(Xml.Name, Xml.Attributes(), run.Xml.Parent!.ElementsBeforeSelf(), leftElement);
            after = new XElement(Xml.Name, Xml.Attributes(), rightElement, run.Xml.Parent.ElementsAfterSelf());
        }
        else
        {
            var (leftElement, rightElement) = run.Split(index - startIndex);
            before = new XElement(Xml.Name, Xml.Attributes(), run.Xml.ElementsBeforeSelf(), leftElement);
            after = new XElement(Xml.Name, Xml.Attributes(), rightElement, run.Xml.ElementsAfterSelf());
        }

        if (!before.Elements().Any()) before = null;
        if (!after.Elements().Any()) after = null;
        return (before, after);
    }

    /// <summary>
    /// Splits a tracked edit (ins/del)
    /// </summary>
    /// <param name="element">Parent element to split</param>
    /// <param name="index">Character index to split on</param>
    /// <returns>Split XElement array</returns>
    private (XElement? leftElement, XElement? rightElement) Split(XElement element, int index)
    {
        // Find the run containing the index
        var (run, startIndex) = FindRunAffectedByEdit(index);

        var (leftElement, rightElement) = run.Split(index-startIndex);

        XElement? splitLeft = new(element.Name, element.Attributes(), run.Xml.ElementsBeforeSelf(), leftElement);
        if (DocumentHelpers.GetTextLength(splitLeft) == 0)
        {
            splitLeft = null;
        }

        XElement? splitRight = new(element.Name, element.Attributes(), rightElement, run.Xml.ElementsAfterSelf());
        if (DocumentHelpers.GetTextLength(splitRight) == 0)
        {
            splitRight = null;
        }

        return (splitLeft, splitRight);
    }

    /// <summary>
    /// Walk all the text runs in the paragraph and find the one containing a specific index.
    /// </summary>
    /// <param name="index">Index to look for</param>
    /// <returns>Run containing index</returns>
    internal (Run, int startIndex) FindRunAffectedByEdit(int index)
    {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));

        int startIndex = 0, total = 0;
        Run? lastRun = null;
        foreach (var run in Runs)
        {
            int size = DocumentHelpers.GetTextLength(run.Xml);
            total += size;
            if (index < total)
            {
                return (run, startIndex);
            }
            startIndex += size;
            lastRun = run;
        }

        if (index == total && lastRun != null)
        {
            return (lastRun, startIndex - DocumentHelpers.GetTextLength(lastRun.Xml));
        }

        throw new ArgumentOutOfRangeException(nameof(index));
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

        Dictionary<int, (string relationId, long chartId)> chartIds = new();
        if (unownedCharts.Any())
        {
            for (var index = 0; index < unownedCharts.Count; index++)
            {
                var chart = unownedCharts[index];
                chartIds.Add((index + 1) * -1, Document.ChartManager.CreateRelationship(chart));
            }
        }

        if (Drawings.Any())
        {
            // Check to see if the .rels file exists and create it if not.
            _ = Document.EnsureRelsPathExists(PackagePart);

            // Fix up pictures
            foreach (var drawing in Drawings)
            {
                drawing.SetOwner(Document, PackagePart, true);

                if (int.TryParse(drawing.ChartRelationId, out var id))
                {
                    if (chartIds.TryGetValue(id, out var relation))
                    {
                        drawing.ChartRelationId = relation.relationId;
                        drawing.Id = relation.chartId;
                    }
                }
            }
        }

        unownedCharts.Clear();
    }

    /// <summary>
    /// Determines equality for paragraphs
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Paragraph? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for paragraphs
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Paragraph);

    /// <summary>
    /// Returns hashcode for this paragraph
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();
}