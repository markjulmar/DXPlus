using DXPlus.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a document paragraph.
    /// </summary>
    [DebuggerDisplay("{xml}")]
    public class Paragraph : InsertBeforeOrAfter, IEquatable<Paragraph>
    {
        /// <summary>
        /// Text runs (r) that make up this paragraph
        /// </summary>
        internal IEnumerable<XElement> Runs => Xml.Elements(Name.Run);

        /// <summary>
        /// Styles in this paragraph
        /// </summary>
        internal List<XElement> Styles { get; } = new List<XElement>();

        /// <summary>
        /// Starting index for this paragraph
        /// </summary>
        internal int StartIndex { get; private set; }

        /// <summary>
        /// End index for this paragraph
        /// </summary>
        internal int EndIndex { get; private set; }

        /// <summary>
        /// Public constructor for the paragraph
        /// </summary>
        public Paragraph() : this(null, Create(null, null), 0)
        {
        }

        /// <summary>
        /// Public constructor for the paragraph
        /// </summary>
        /// <param name="text"></param>
        public Paragraph(string text) : this (null, Create(text, null), 0)
        {
        }

        /// <summary>
        /// Public constructor for the paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <param name="formatting"></param>
        public Paragraph(string text, Formatting formatting) : this(null, Create(text, formatting), 0)
        {
        }

        /// <summary>
        /// Constructor for the paragraph
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">XML for the paragraph</param>
        /// <param name="startIndex">Starting position in the doc</param>
        /// <param name="parentContainerType">Container parent type</param>
        internal Paragraph(IDocument document, XElement xml, int startIndex, ContainerType parentContainerType = ContainerType.None)
            : base(document, xml)
        {
            ParentContainerType = parentContainerType;
            StartIndex = startIndex;
            EndIndex = startIndex + GetElementTextLength(Xml);
        }

        /// <summary>
        /// Gets or set this Paragraphs text alignment.
        /// </summary>
        public Alignment Alignment
        {
            get => GetDefaultFormatting().Alignment;
            set => GetDefaultFormatting(true).Alignment = value;
        }

        /// <summary>
        /// Returns the run properties, or if we don't have any, the paragraph properties.
        /// </summary>
        /// <returns></returns>
        internal Formatting GetFormatting(bool create = false)
        {
            XElement rPr;

            // If we don't have any text runs yet, consider this to be the default properties.
            if (!Runs.Any())
            {
                if (create)
                {
                    XElement pPr = Xml.Element(Name.ParagraphProperties);
                    if (pPr == null)
                    {
                        pPr = new XElement(Name.ParagraphProperties);
                        Xml.AddFirst(pPr);
                    }
                    rPr = pPr.GetRunProps(true);
                }
                else
                {
                    rPr = Xml.Element(Name.ParagraphProperties, Name.RunProperties);
                }
            }
            else
            {
                // Otherwise get/create on the last text run. If we don't have any text
                // runs, use the last run available.
                List<XElement> runs = Runs.Reverse().ToList();
                rPr = (runs.Find(r => r.Element(Name.Text) != null) ?? runs[0]).GetRunProps(create);
            }

            return new Formatting(rPr);
        }

        /// <summary>
        /// Returns whether this paragraph is marked as BOLD
        /// </summary>
        public bool Bold
        {
            get => GetFormatting().Bold;
            set => GetFormatting(true).Bold = value;
        }

        /// <summary>
        /// Change the italic state of this paragraph
        /// </summary>
        public bool Italic
        {
            get => GetFormatting().Italic;
            set => GetFormatting(true).Italic = value;
        }

        /// <summary>
        /// Change the paragraph to be small caps, capitals or none.
        /// </summary>
        public CapsStyle CapsStyle
        {
            get => GetFormatting().CapsStyle;
            set => GetFormatting(true).CapsStyle = value;
        }

        /// <summary>
        /// Returns the applied text color, or None for default.
        /// </summary>
        public Color Color
        {
            get => GetFormatting().Color;
            set => GetFormatting(true).Color = value;
        }

        /// <summary>
        /// Change the culture of the given paragraph.
        /// </summary>
        public CultureInfo Culture
        {
            get => GetFormatting().Culture;
            set => GetFormatting(true).Culture = value;
        }

        /// <summary>
        /// Change the font for the paragraph
        /// </summary>
        public FontFamily Font
        {
            get => GetFormatting().Font;
            set => GetFormatting(true).Font = value;
        }

        /// <summary>
        /// Get or set the font size of this paragraph
        /// </summary>
        public double? FontSize
        {
            get => GetFormatting().FontSize;
            set => GetFormatting(true).FontSize = value;
        }

        /// <summary>
        /// Gets or Sets the Direction of content in this Paragraph.
        /// </summary>
        public Direction Direction
        {
            get => Xml.Element(Name.ParagraphProperties, Name.RTL) == null ? Direction.LeftToRight : Direction.RightToLeft;

            set
            {
                if (value == Direction.RightToLeft)
                {
                    Xml.GetOrCreateElement(Name.ParagraphProperties, Name.RTL);
                }
                else
                {
                    Xml.Element(Name.ParagraphProperties, Name.RTL)?.Remove();
                }
            }
        }

        /// <summary>
        /// True if this paragraph is hidden.
        /// </summary>
        public bool IsHidden
        {
            get => GetFormatting().IsHidden;
            set => GetFormatting(true).IsHidden = value;
        }

        /// <summary>
        /// Gets or sets the highlight on this paragraph
        /// </summary>
        public Highlight Highlight
        {
            get => GetFormatting().Highlight;
            set => GetFormatting(true).Highlight = value;
        }

        /// <summary>
        /// Set the kerning for the paragraph
        /// </summary>
        public int? Kerning
        {
            get => GetFormatting().Kerning;
            set => GetFormatting(true).Kerning = value;
        }

        /// <summary>
        /// Applied effect on the paragraph
        /// </summary>
        public Effect Effect
        {
            get => GetFormatting().Effect;
            set => GetFormatting(true).Effect = value;
        }

        /// <summary>
        /// Returns a list of DocProperty elements in this document.
        /// </summary>
        public IEnumerable<DocProperty> DocumentProperties
        {
            get
            {
                if (Document == null)
                {
                    throw new InvalidOperationException("Cannot use document properties without a document owner.");
                }

                return Xml.Descendants(Name.SimpleField)
                    .Select(el => new DocProperty(Document, el));
            }
        }

        ///<summary>
        /// Returns table following the paragraph. Null if the following element isn't table.
        ///</summary>
        public Table FollowingTable { get; internal set; }

        /// <summary>
        /// Set the left indentation in 1/20th pt for this Paragraph.
        /// </summary>
        public double LeftIndent
        {
            get => GetDefaultFormatting().IndentationLeft;
            set => GetDefaultFormatting(true).IndentationLeft = value;
        }

        /// <summary>
        /// Set the right indentation in 1/20th pt for this Paragraph.
        /// </summary>
        public double RightIndent

        {
            get => GetDefaultFormatting().IndentationRight;
            set => GetDefaultFormatting(true).IndentationRight = value;
        }

        /// <summary>
        /// Get or set the indentation of the first line of this Paragraph.
        /// </summary>
        public double FirstLineIndent
        {
            get => GetDefaultFormatting().IndentationFirstLine;
            set => GetDefaultFormatting(true).IndentationFirstLine = value;
        }

        /// <summary>
        /// Get or set the indentation of all but the first line of this Paragraph.
        /// </summary>
        public double HangingIndent
        {
            get => GetDefaultFormatting().IndentationHanging;
            set => GetDefaultFormatting(true).IndentationHanging = value;
        }

        /// <summary>
        /// Paragraph formatting
        /// </summary>
        private ParagraphProperties GetDefaultFormatting(bool create = false)
        {
            XElement pPr = Xml.Element(Name.ParagraphProperties);
            if (create && pPr == null)
            {
                pPr = new XElement(Name.ParagraphProperties);
                Xml.AddFirst(pPr);
            }

            return new ParagraphProperties(pPr);
        }

        /// <summary>
        /// True to keep with the next element on the page.
        /// </summary>
        public bool KeepWithNext
        {
            get => GetDefaultFormatting().KeepWithNext;
            set => GetDefaultFormatting(true).KeepWithNext = value;
        }

        /// <summary>
        /// Keep lines together on the page
        /// </summary>
        public bool KeepLinesTogether
        {
            get => GetDefaultFormatting().KeepLinesTogether;
            set => GetDefaultFormatting(true).KeepLinesTogether = value;
        }

        /// <summary>
        /// Container owner type.
        /// </summary>
        public ContainerType ParentContainerType { get; set; }

        /// <summary>
        /// Specifies the amount by which each character shall be expanded or when the character is rendered in the document.
        /// This property stretches or compresses each character in the run.
        /// </summary>
        public int? ExpansionScale
        {
            get => GetFormatting().ExpansionScale;
            set => GetFormatting(true).ExpansionScale = value;
        }

        /// <summary>
        /// Set the spacing between lines in this paragraph
        /// </summary>
        public double? LineSpacing
        {
            get => GetDefaultFormatting().LineSpacing;
            set => GetDefaultFormatting(true).LineSpacing = value;
        }

        /// <summary>
        /// Set the spacing after lines in this paragraph
        /// </summary>
        public double? LineSpacingAfter
        {
            get => GetDefaultFormatting().LineSpacingAfter;
            set => GetDefaultFormatting(true).LineSpacingAfter = value;
        }

        /// <summary>
        /// Set the spacing before lines in this paragraph
        /// </summary>
        public double? LineSpacingBefore
        {
            get => GetDefaultFormatting().LineSpacingBefore;
            set => GetDefaultFormatting(true).LineSpacingBefore = value;
        }

        /// <summary>
        /// Specifies the amount by which text shall be raised or lowered for this run in relation to the default
        /// baseline of the surrounding non-positioned text. This allows the text to be repositioned without
        /// altering the font size of the contents. This is measured in pts.
        /// </summary>
        public double? Position
        {
            get => GetFormatting().Position;
            set => GetFormatting(true).Position = value;
        }

        /// <summary>
        /// Returns a list of all Pictures in a Paragraph.
        /// </summary>
        public List<Picture> Pictures => (
                    from p in Xml.LocalNameDescendants("drawing")
                    let id = p.FirstLocalNameDescendant("blip").AttributeValue(Namespace.RelatedDoc + "embed")
                    where id != null
                    let img = new Image(Document, Document.PackagePart.GetRelationship(id))
                    select new Picture(Document, p, img)
                ).Union(
                    from p in Xml.LocalNameDescendants("pict")
                    let id = p.FirstLocalNameDescendant("imagedata").AttributeValue(Namespace.RelatedDoc + "id")
                    where id != null
                    let img = new Image(Document, Document.PackagePart.GetRelationship(id))
                    select new Picture(Document, p, img)
                ).ToList();

        /// <summary>
        /// Set the paragraph to subscript. Note this is mutually exclusive with Superscript
        /// </summary>
        public bool Subscript
        {
            get => GetFormatting().Subscript;
            set => GetFormatting(true).Subscript = value;
        }

        /// <summary>
        /// Set the paragraph to Superscript. Note this is mutually exclusive with Subscript.
        /// </summary>
        public bool Superscript
        {
            get => GetFormatting().Superscript;
            set => GetFormatting(true).Superscript = value;
        }

        ///<summary>
        /// The style name of the paragraph.
        ///</summary>
        public string StyleName
        {
            get => GetDefaultFormatting().StyleName;
            set => GetDefaultFormatting(true).StyleName = value;
        }

        /// <summary>
        /// Get or set the underline style for this paragraph
        /// </summary>
        public UnderlineStyle UnderlineStyle
        {
            get => GetFormatting().UnderlineStyle;
            set => GetFormatting(true).UnderlineStyle = value;
        }

        /// <summary>
        /// Get or set the emphasis on the last run
        /// </summary>
        public Emphasis Emphasis
        {
            get => GetFormatting().Emphasis;
            set => GetFormatting(true).Emphasis = value;
        }

        /// <summary>
        /// Gets the text value of this Paragraph.
        /// </summary>
        public string Text => HelperFunctions.GetText(Xml);

        /// <summary>
        /// Append text to this Paragraph.
        /// </summary>
        /// <param name="text">The text to append.</param>
        /// <returns>This Paragraph with the new text appended.</returns>
        public Paragraph Append(string text)
        {
            Xml.Add(HelperFunctions.FormatInput(text, null));
            return this;
        }

        /// <summary>
        /// Appends a new bookmark to the paragraph
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <returns>This paragraph</returns>
        public Paragraph AppendBookmark(string bookmarkName)
        {
            Xml.Add(new XElement(Name.BookmarkStart,
                        new XAttribute(Name.Id, 0),
                        new XAttribute(Name.NameId, bookmarkName)));

            Xml.Add(new XElement(Name.BookmarkEnd,
                        new XAttribute(Name.Id, 0),
                        new XAttribute(Name.NameId, bookmarkName)));

            return this;
        }

        /// <summary>
        /// Add an equation to a document.
        /// </summary>
        /// <param name="equation">The Equation to append.</param>
        /// <returns>The Paragraph with the Equation now appended.</returns>
        public Paragraph AppendEquation(string equation)
        {
            // Create equation element
            XElement oMathPara = new XElement(Name.MathParagraph,
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
        /// This function inserts a hyperlink into a Paragraph at a specified character index.
        /// </summary>
        /// <param name="hyperlink">The hyperlink to insert.</param>
        /// <param name="charIndex">The character index in the owning paragraph to insert at.</param>
        /// <returns>The Paragraph with the Hyperlink inserted at the specified index.</returns>
        public Paragraph InsertHyperlink(Hyperlink hyperlink, int charIndex = 0)
        {
            // Set the package and document relationship.
            hyperlink.Document = Document;
            hyperlink.PackagePart = PackagePart;

            // Check to see if the rels file exists and create it if not.
            _ = Document.EnsureRelsPathExists(PackagePart);

            // Check to see if a rel for this Hyperlink exists, create it if not.
            _ = hyperlink.GetOrCreateRelationship();

            if (charIndex == 0)
            {
                // Add this hyperlink as the last element.
                Xml.AddFirst(hyperlink.Xml);
            }
            else
            {
                // Get the first run effected by this Insert
                Run run = GetFirstRunAffectedByEdit(charIndex);
                if (run == null)
                {
                    // Add this hyperlink as the last element.
                    Xml.Add(hyperlink.Xml);
                }
                else
                {
                    // Split this run at the point you want to insert
                    XElement[] splitRun = run.SplitRun(charIndex);

                    // Replace the original run.
                    run.Xml.ReplaceWith(splitRun[0], hyperlink.Xml, splitRun[1]);
                }
            }

            return this;
        }

        /// <summary>
        /// Returns a list of Hyperlinks in this Paragraph.
        /// </summary>
        public List<Hyperlink> Hyperlinks => Hyperlink.Enumerate(this).ToList();

        /// <summary>
        /// Append a hyperlink to a Paragraph.
        /// </summary>
        /// <param name="hyperlink">The hyperlink to append.</param>
        /// <returns>The Paragraph with the hyperlink appended.</returns>
        public Paragraph Append(Hyperlink hyperlink)
        {
            if (Document == null)
            {
                throw new InvalidOperationException("Cannot add hyperlinks to paragraphs with no document owner.");
            }

            // Ensure the owner document has the hyperlink styles.
            Document.AddHyperlinkStyleIfNotPresent();

            // Set the document/package for the hyperlink to be owned by the document
            hyperlink.Document = Document;
            hyperlink.PackagePart = PackagePart;

            // Check to see if the rels file exists and create it if not.
            _ = Document.EnsureRelsPathExists(PackagePart);
            _ = hyperlink.GetOrCreateRelationship();

            Xml.Add(hyperlink.Xml);

            return this;
        }

        /// <summary>
        /// Append a PageCount place holder onto the end of a Paragraph.
        /// </summary>
        /// <param name="format">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        public void AppendPageCount(PageNumberFormat format)
        {
            AddPageNumberInfo(format, "numPages");
        }

        /// <summary>
        /// Append a PageNumber place holder onto the end of a Paragraph.
        /// </summary>
        /// <param name="format">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        public void AppendPageNumber(PageNumberFormat format)
        {
            AddPageNumberInfo(format, "page");
        }

        /// <summary>
        /// Insert a PageCount place holder into a Paragraph.
        /// This place holder should only be inserted into a Header or Footer Paragraph.
        /// Word will not automatically update this field if it is inserted into a document level Paragraph.
        /// </summary>
        /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        /// <param name="index">The text index to insert this PageCount place holder at.</param>
        public void InsertPageCount(PageNumberFormat pnf, int index = 0)
        {
            AddPageNumberInfo(pnf, "numPages", index);
        }

        /// <summary>
        /// Insert a PageNumber place holder into a Paragraph.
        /// This place holder should only be inserted into a Header or Footer Paragraph.
        /// Word will not automatically update this field if it is inserted into a document level Paragraph.
        /// </summary>
        /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        /// <param name="index">The text index to insert this PageNumber place holder at.</param>
        public void InsertPageNumber(PageNumberFormat pnf, int index = 0)
        {
            AddPageNumberInfo(pnf, "page", index);
        }

        /// <summary>
        /// Internal method to populate page numbers or page counts.
        /// </summary>
        /// <param name="format">Page number format</param>
        /// <param name="type">Numbers or Counts</param>
        /// <param name="index">Position to insert</param>
        private void AddPageNumberInfo(PageNumberFormat format, string type, int? index = null)
        {
            XElement fldSimple = new XElement(Name.SimpleField);

            fldSimple.Add(format == PageNumberFormat.Normal
                ? new XAttribute(Namespace.Main + "instr", $@" {type.ToUpper()}   \* MERGEFORMAT ")
                : new XAttribute(Namespace.Main + "instr", $@" {type.ToUpper()}  \* ROMAN  \* MERGEFORMAT "));

            XElement content = XElement.Parse(
                @"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof />
                   </w:rPr>
                   <w:t>1</w:t>
               </w:r>"
            );

            if (index == null || index.Value < 0)
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
                Run r = GetFirstRunAffectedByEdit(index.Value);
                XElement[] splitEdit = SplitEdit(r.Xml, index.Value, EditType.Insert);
                r.Xml.ReplaceWith(splitEdit[0], fldSimple, splitEdit[1]);
            }
        }

        /// <summary>
        /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
        /// </summary>
        /// <param name="picture">The Picture to append.</param>
        /// <returns>The Paragraph with the Picture now appended.</returns>
        public Paragraph Append(Picture picture)
        {
            if (Document == null)
                throw new ArgumentException("Cannot add pictures without a document owner.");

            // Check to see if the .rels file exists and create it if not.
            _ = Document.EnsureRelsPathExists(PackagePart);

            // Check to see if a rel for this Picture exists, create it if not.
            picture.PackagePart = this.PackagePart;
            picture.Id = picture.GetOrCreateRelationship();

            // Add a new run with the given drawing to the paragraph.
            var run = new XElement(Name.Run,
                new XElement(Name.RunProperties,
                    new XElement(Namespace.Main + "noProof")),
                picture.Xml);
            Xml.Add(run);

            return this;
        }

        /// <summary>
        /// Find all instances of a string in this paragraph and return their indexes in a List.
        /// </summary>
        /// <param name="text">The string to find</param>
        /// <param name="ignoreCase">True to ignore case in the search</param>
        /// <returns>A list of indexes.</returns>
        public IEnumerable<int> FindAll(string text, bool ignoreCase)
        {
            MatchCollection mc = Regex.Matches(Text, Regex.Escape(text), ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
            return mc.Select(m => m.Index);
        }

        /// <summary>
        ///  Find all unique instances of the given Regex Pattern
        /// </summary>
        /// <param name="regex">Regex to match</param>
        /// <returns>Index and matched text</returns>
        public IEnumerable<(int index, string text)> FindPattern(Regex regex)
        {
            MatchCollection mc = regex.Matches(Text);
            return mc.Select(m => (index: m.Index, text: m.Value));
        }

        /// <summary>
        /// Retrieve all the bookmarks in this paragraph
        /// </summary>
        /// <returns>Enumerable of bookmark objects</returns>
        public IEnumerable<Bookmark> GetBookmarks()
        {
            return Xml.Descendants(Name.BookmarkStart)
                        .Select(x => x.Attribute(Name.NameId))
                        .Where(x => x != null)
                        .Select(x => new Bookmark(x.Value, this));
        }

        /// <summary>
        /// Insert a text block at a bookmark
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <param name="toInsert">Text to insert</param>
        public bool InsertAtBookmark(string bookmarkName, string toInsert)
        {
            XElement bookmark = Xml.Descendants(Name.BookmarkStart)
                    .SingleOrDefault(x => x.AttributeValue(Name.NameId) == bookmarkName);
            if (bookmark != null)
            {
                List<XElement> run = HelperFunctions.FormatInput(toInsert, null);
                bookmark.AddBeforeSelf(run);
                Document?.RenumberIds();
                return true;
            }
            return false;
        }

        /// <summary>
        /// Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
        /// </summary>
        /// <param name="cp">The custom property to display.</param>
        /// <param name="formatting">The formatting to use for this text.</param>
        public DocProperty AddDocumentProperty(CustomProperty cp, Formatting formatting = null)
        {
            if (Document == null)
            {
                throw new InvalidOperationException("Cannot add document properties without a document owner.");
            }

            XElement xml = new XElement(Name.SimpleField,
                new XAttribute(Namespace.Main + "instr", $@"DOCPROPERTY {cp.Name} \* MERGEFORMAT"),
                    new XElement(Name.Run,
                        new XElement(Name.Text, formatting?.Xml, cp.Value))
            );

            Xml.Add(xml);

            return new DocProperty(Document, xml);
        }

        /// <summary>
        /// Insert a Picture into a Paragraph at the given text index.
        /// If not index is provided defaults to 0.
        /// </summary>
        /// <param name="picture">The Picture to insert.</param>
        /// <param name="index">The text index to insert at.</param>
        /// <returns>The modified Paragraph.</returns>
        public Paragraph Insert(Picture picture, int index = 0)
        {
            if (Document == null)
            {
                throw new InvalidOperationException("Cannot add pictures without a document owner.");
            }

            // Check to see if the rels file exists and create it if not.
            _ = Document.EnsureRelsPathExists(PackagePart);

            // Check to see if a rel for this Picture exists, create it if not.
            picture.PackagePart = this.PackagePart;
            picture.Id = picture.GetOrCreateRelationship();

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
                Run run = GetFirstRunAffectedByEdit(index);
                if (run == null)
                {
                    // Add this picture as the last element.
                    Xml.Add(xml);
                }
                else
                {
                    // Split this run at the point you want to insert
                    XElement[] splitRun = run.SplitRun(index);

                    // Replace the original run.
                    run.Xml.ReplaceWith(splitRun[0], xml, splitRun[1]);
                }
            }

            return this;
        }

        /// <summary>
        /// Inserts a string into a Paragraph with the specified formatting.
        /// </summary>
        public void InsertText(string value, Formatting formatting = null)
        {
            Xml.Add(HelperFunctions.FormatInput(value, formatting?.Xml));
            Document?.RenumberIds();
        }

        /// <summary>
        /// Replaces any existing text runs in this paragraph with the specified text.
        /// </summary>
        /// <param name="value">Text value</param>
        /// <param name="formatting">Formatting to apply (null for default)</param>
        public void SetText(string value, Formatting formatting = null)
        {
            // Remove all runs from this paragraph.
            Xml.Descendants(Name.Run).Remove();

            if (!string.IsNullOrEmpty(value))
            {
                // Add the new run.
                Xml.Add(HelperFunctions.FormatInput(value, formatting?.Xml));
                // Renumber any insert/delete markers
                Document?.RenumberIds();
            }
        }

        /// <summary>
        /// Inserts a string into the Paragraph with the specified formatting at the given index.
        /// </summary>
        /// <param name="index">The index position of the insertion.</param>
        /// <param name="value">The System.String to insert.</param>
        /// <param name="formatting">The text formatting.</param>
        public void InsertText(int index, string value, Formatting formatting = null)
        {
            // Get the first run effected by this Insert
            Run run = GetFirstRunAffectedByEdit(index);
            if (run == null)
            {
                Xml.Add(HelperFunctions.FormatInput(value, formatting?.Xml));
            }
            else
            {
                object insert = HelperFunctions.FormatInput(value, formatting?.Xml);
                XElement parentElement = run.Xml.Parent;
                if (parentElement == null)
                {
                    throw new InvalidOperationException("Orphaned run not connected to paragraph.");
                }

                switch (parentElement.Name.LocalName)
                {
                    case "ins":
                    case "del":
                        XElement[] splitEdit = SplitEdit(parentElement, index, EditType.Insert);
                        parentElement.ReplaceWith(splitEdit[0], insert, splitEdit[1]);
                        break;

                    default:
                        XElement[] splitRun = run.SplitRun(index);
                        run.Xml.ReplaceWith(splitRun[0], insert, splitRun[1]);
                        break;
                }
            }

            Document?.RenumberIds();
        }

        /// <summary>
        /// Remove this Paragraph from the document.
        /// </summary>
        public void Remove()
        {
            // If this is the only Paragraph in the Cell then we cannot remove it.
            if (Xml.Parent?.Name.LocalName == "tc"
                && Xml.Parent.Elements(Name.Paragraph).Count() == 1)
            {
                Xml.Value = string.Empty;
            }
            else
            {
                // Remove this paragraph from the document
                Xml.Remove();
                Xml = null;
            }
        }

        /// <summary>
        /// Remove the Hyperlink at the provided index. The first hyperlink is at index 0.
        /// Using a negative index or an index greater than the index of the last hyperlink will cause an ArgumentOutOfRangeException() to be thrown.
        /// </summary>
        /// <param name="index">The index of the hyperlink to be removed.</param>
        public void RemoveHyperlink(int index)
        {
            int count = 0;
            if (index < 0 || !RemoveHyperlinkRecursive(Xml, index, ref count))
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
        }

        /// <summary>
        /// Removes characters from a DXPlus.DocX.Paragraph.
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
                Run run = GetFirstRunAffectedByEdit(index, EditType.Delete);

                // The parent of this Run
                XElement parentElement = run.Xml.Parent;
                switch (parentElement?.Name.LocalName)
                {
                    case "ins":
                    case "del":
                        {
                            XElement[] splitEditBefore = SplitEdit(parentElement, index, EditType.Delete);
                            int take = count - processed;
                            XElement[] splitEditAfter = SplitEdit(parentElement, index + take, EditType.Delete);
                            XElement middle = SplitEdit(splitEditBefore[1], index + take, EditType.Delete)[1];
                            processed += GetElementTextLength(middle);
                            parentElement.ReplaceWith(splitEditBefore[0], null, splitEditAfter[1]);
                        }
                        break;

                    default:
                        if (GetElementTextLength(run.Xml) > 0)
                        {
                            XElement[] splitRunBefore = run.SplitRun(index, EditType.Delete);
                            int min = Math.Min(index + (count - processed), run.EndIndex);
                            XElement[] splitRunAfter = run.SplitRun(min, EditType.Delete);
                            XElement middle = new Run(splitRunBefore[1], run.StartIndex + GetElementTextLength(splitRunBefore[0])).SplitRun(min, EditType.Delete)[0];
                            processed += GetElementTextLength(middle);
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
                    && GetElementTextLength(parentElement) == 0
                    && parentElement.Parent != null
                    && parentElement.Parent.Name.LocalName != "tc"
                    && parentElement.Parent.Elements(Name.Paragraph).Any()
                    && !parentElement.Descendants(Namespace.Main + "drawing").Any())
                {
                    parentElement.Remove();
                }
            }
            while (processed < count);

            Document?.RenumberIds();
        }

        /// <summary>
        /// Replaces all occurrences of a specified System.String in this instance, with another specified System.String.
        /// </summary>
        /// <param name="oldValue">Text to be replaced.</param>
        /// <param name="newValue">Text to replace all occurrences of oldValue.</param>
        /// <param name="options">RegexOption enumeration options.</param>
        /// <param name="newFormatting">The formatting to apply to the text being inserted.</param>
        /// <param name="matchFormatting">The formatting that the text must match in order to be replaced.</param>
        /// <param name="fo">How should formatting be matched?</param>
        /// <param name="escapeRegEx">True if the oldValue needs to be escaped, otherwise false. If it represents a valid RegEx pattern this should be false.</param>
        /// <param name="useRegExSubstitutions">True if RegEx-like replace should be performed, i.e. if newValue contains RegEx substitutions. Does not perform named-group substitutions (only numbered groups).</param>
        public void ReplaceText(string oldValue, string newValue, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch, bool escapeRegEx = true, bool useRegExSubstitutions = false)
        {
            string tText = Text;
            MatchCollection mc = Regex.Matches(tText, escapeRegEx ? Regex.Escape(oldValue) : oldValue, options);

            // Loop through the matches in reverse order
            foreach (Match m in mc.Reverse())
            {
                // Assume the formatting matches until proven otherwise.
                bool formattingMatch = true;

                // Does the user want to match formatting?
                if (matchFormatting != null)
                {
                    // The number of characters processed so far
                    int processed = 0;

                    do
                    {
                        // Get the next run effected
                        Run run = GetFirstRunAffectedByEdit(m.Index + processed);

                        // Get this runs properties
                        XElement rPr = run.Xml.GetRunProps(false) ?? new Formatting().Xml;

                        // Make sure that every formatting element in f.xml is also in this run,
                        // if this is not true, then their formatting does not match.
                        if (!HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, rPr, fo))
                        {
                            formattingMatch = false;
                            break;
                        }

                        // We have processed some characters, so update the counter.
                        processed += run.Text.Length;
                    } while (processed < m.Length);
                }

                // If the formatting matches, do the replace.
                if (formattingMatch)
                {
                    string repl = newValue;
                    if (useRegExSubstitutions && !string.IsNullOrEmpty(repl))
                    {
                        repl = repl.Replace("$&", m.Value);
                        if (m.Groups.Count > 0)
                        {
                            int lastCaptureIndex = 0;
                            for (int k = 0; k < m.Groups.Count; k++)
                            {
                                Group g = m.Groups[k];
                                if (g.Value.Length == 0)
                                {
                                    continue;
                                }

                                repl = repl.Replace("$" + k, g.Value);
                                lastCaptureIndex = k;
                            }
                            repl = repl.Replace("$+", m.Groups[lastCaptureIndex].Value);
                        }
                        if (m.Index > 0)
                        {
                            repl = repl.Replace("$`", tText.Substring(0, m.Index));
                        }
                        if (m.Index + m.Length < tText.Length)
                        {
                            repl = repl.Replace("$'", tText.Substring(m.Index + m.Length));
                        }
                        repl = repl.Replace("$_", tText);
                        repl = repl.Replace("$$", "$");
                    }
                    if (!string.IsNullOrEmpty(repl))
                    {
                        InsertText(m.Index + m.Length, repl, newFormatting);
                    }

                    if (m.Length > 0)
                    {
                        RemoveText(m.Index, m.Length);
                    }
                }
            }
        }

        /// <summary>
        /// Find pattern regex must return a group match.
        /// </summary>
        /// <param name="findPattern">Regex pattern that must include one group match. ie (.*)</param>
        /// <param name="regexMatchHandler">A func that accepts the matching find grouping text and returns a replacement value</param>
        /// <param name="options"></param>
        /// <param name="newFormatting"></param>
        /// <param name="matchFormatting"></param>
        /// <param name="fo"></param>
        public void ReplaceText(string findPattern, Func<string, string> regexMatchHandler,
            RegexOptions options = RegexOptions.None, Formatting newFormatting = null,
            Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch)
        {
            MatchCollection matchCollection = Regex.Matches(Text, findPattern, options);

            // Loop through the matches in reverse order
            foreach (Match match in matchCollection.Reverse())
            {
                // Assume the formatting matches until proven otherwise.
                bool formattingMatch = true;

                // Does the user want to match formatting?
                if (matchFormatting != null)
                {
                    // The number of characters processed so far
                    int processed = 0;

                    do
                    {
                        // Get the next run effected
                        Run run = GetFirstRunAffectedByEdit(match.Index + processed);

                        // Get this runs properties
                        XElement rPr = run.Xml.GetRunProps(false) ?? new Formatting().Xml;
                        if (!HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, rPr, fo))
                        {
                            formattingMatch = false;
                            break;
                        }

                        // We have processed some characters, so update the counter.
                        processed += run.Text.Length;
                    } while (processed < match.Length);
                }

                // If the formatting matches, do the replace.
                if (formattingMatch)
                {
                    string newValue = regexMatchHandler.Invoke(match.Groups[1].Value);
                    InsertText(match.Index + match.Value.Length, newValue, newFormatting);
                    RemoveText(match.Index, match.Value.Length);
                }
            }
        }

        /// <summary>
        /// Specifies that the text in this paragraph should be displayed with a single or double-line
        /// strikethrough
        /// </summary>
        public Strikethrough StrikeThrough
        {
            get => GetFormatting().StrikeThrough;
            set => GetFormatting(true).StrikeThrough = value;
        }

        /// <summary>
        /// Unique id for this paragraph
        /// </summary>
        public string Id
        {
            get => Xml.AttributeValue(Name.ParagraphId, null);
            set => Xml.SetAttributeValue(Name.ParagraphId, string.IsNullOrEmpty(value) ? null : value);
        }

        /// <summary>
        /// Append text to this Paragraph and then underline it using a color.
        /// </summary>
        /// <param name="underlineColor">The underline color to use, if no underline is set, a single line will be used.</param>
        /// <returns>This Paragraph with the last appended text underlined in a color.</returns>
        public Paragraph UnderlineColor(Color underlineColor)
        {
            foreach (XElement run in Runs)
            {
                _ = new Formatting(run.GetRunProps(true)) { UnderlineColor = underlineColor };
            }

            return this;
        }

        /// <summary>
        /// Retrieve the text length of the passed element
        /// </summary>
        /// <param name="textElement"></param>
        /// <returns></returns>
        internal static int GetElementTextLength(XElement textElement)
        {
            int count = 0;
            if (textElement != null)
            {
                foreach (XElement el in textElement.Descendants())
                {
                    switch (el.Name.LocalName)
                    {
                        case "tab":
                            if (el.Parent?.Name.LocalName != "tabs")
                            {
                                goto case "br";
                            }

                            break;

                        case "br":
                            count++;
                            break;

                        case "t":
                        case "delText":
                            count += el.Value.Length;
                            break;
                    }
                }
            }
            return count;
        }

        /// <summary>
        /// Walk all the text runs in the paragraph and find the one containing a specific index.
        /// </summary>
        /// <param name="index">Index to look for</param>
        /// <param name="editType">Type of edit being performed (insert or delete)</param>
        /// <returns>Run containing index</returns>
        internal Run GetFirstRunAffectedByEdit(int index, EditType editType = EditType.Insert)
        {
            int len = HelperFunctions.GetText(Xml).Length;
            if (index < 0 || (editType == EditType.Insert && index > len) || (editType == EditType.Delete && index >= len))
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            int count = 0;
            Run theOne = null;
            GetFirstRunAffectedByEditRecursive(Xml, index, ref count, ref theOne, editType);

            return theOne;
        }

        /// <summary>
        /// Recursive method to identify a text run from a starting element and index.
        /// </summary>
        /// <param name="el">Element to search</param>
        /// <param name="index">Index to look for</param>
        /// <param name="count">Total searched</param>
        /// <param name="theOne">The located text run</param>
        /// <param name="editType">Type of edit being performed (insert or delete)</param>
        private void GetFirstRunAffectedByEditRecursive(XElement el, int index, ref int count, ref Run theOne, EditType editType)
        {
            count += HelperFunctions.GetSize(el);

            // If the EditType is deletion then we must return the next blah
            if (count > 0 && ((editType == EditType.Delete && count > index) || (editType == EditType.Insert && count >= index)))
            {
                // Correct the index
                count -= el.ElementsBeforeSelf().Sum(HelperFunctions.GetSize);
                count -= HelperFunctions.GetSize(el);
                count = Math.Max(0, count);

                // We have found the element, now find the run it belongs to.
                while (el != null && el.Name.LocalName != "r")
                {
                    el = el.Parent;
                }

                if (el == null)
                {
                    return;
                }

                theOne = new Run(el, count);
            }
            else if (el.HasElements)
            {
                foreach (XElement e in el.Elements())
                {
                    if (theOne == null)
                    {
                        GetFirstRunAffectedByEditRecursive(e, index, ref count, ref theOne, editType);
                    }
                }
            }
        }

        /// <summary>
        /// Removes a specific hyperlink by index through a recursive search
        /// </summary>
        /// <param name="element">Element to search</param>
        /// <param name="index">Index to look for</param>
        /// <param name="count"># of hyperlinks found so far</param>
        /// <returns>True when hyperlink is removed</returns>
        internal bool RemoveHyperlinkRecursive(XElement element, int index, ref int count)
        {
            if (element.Name.LocalName.Equals("hyperlink", StringComparison.CurrentCultureIgnoreCase))
            {
                // Count the number of hyperlinks we've found so far. When we hit the
                // index, that's the one we want to remove.
                if (count == index)
                {
                    element.Remove();
                    return true;
                }
                count++;
            }

            foreach (XElement e in element.Elements())
            {
                if (RemoveHyperlinkRecursive(e, index, ref count))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Splits a tracked edit (ins/del)
        /// </summary>
        /// <param name="element">Parent element to split</param>
        /// <param name="index">Character index to split on</param>
        /// <param name="type">Type of edit being performed (insert/delete)</param>
        /// <returns>Split XElement array</returns>
        internal XElement[] SplitEdit(XElement element, int index, EditType type)
        {
            // Find the run containing the index
            Run run = GetFirstRunAffectedByEdit(index, type);
            XElement[] splitRun = run.SplitRun(index, type);

            XElement splitLeft = new XElement(element.Name, element.Attributes(), run.Xml.ElementsBeforeSelf(), splitRun[0]);
            if (GetElementTextLength(splitLeft) == 0)
            {
                splitLeft = null;
            }

            XElement splitRight = new XElement(element.Name, element.Attributes(), splitRun[1], run.Xml.ElementsAfterSelf());
            if (GetElementTextLength(splitRight) == 0)
            {
                splitRight = null;
            }

            return new[] { splitLeft, splitRight };
        }

        /// <summary>
        /// Provides value equality for the paragraph.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(Paragraph other)
        {
            if (other is null)
            {
                return false;
            }

            if (ReferenceEquals(this, other))
            {
                return true;
            }

            return Text == other.Text
                   && Id == other.Id
                   && StartIndex == other.StartIndex
                   && EndIndex == other.EndIndex
                   && ParentContainerType == other.ParentContainerType;
        }

        protected override void OnAddedToContainer(Container container)
        {
            base.OnAddedToContainer(container);
            if (container != null)
            {
                this.ParentContainerType = container.GetType().Name switch
                {
                    nameof(Table) => ContainerType.Table,
                    nameof(Section) => ContainerType.Section,
                    nameof(Cell) => ContainerType.Cell,
                    nameof(Header) => ContainerType.Header,
                    nameof(Footer) => ContainerType.Footer,
                    nameof(Paragraph) => ContainerType.Paragraph,
                    nameof(DocX) => ContainerType.Body,
                    _ => ContainerType.None
                };
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
            EndIndex = index + GetElementTextLength(Xml);
        }

        /// <summary>
        /// Create a new paragraph from some text
        /// </summary>
        /// <param name="text"></param>
        /// <param name="formatting"></param>
        /// <returns>New paragraph</returns>
        internal static XElement Create(string text, Formatting formatting)
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
            return new Paragraph
            {
                Xml = otherParagraph.Xml.Clone(), Id = HelperFunctions.GenerateHexId()
            };
        }
    }
}