using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a document paragraph.
    /// </summary>
    public class Paragraph : InsertBeforeOrAfter
    {
        /// <summary>
        /// Text runs (r) that make up this paragraph
        /// </summary>
        internal List<XElement> Runs { get; set; }

        /// <summary>
        /// Styles in this paragraph
        /// </summary>
        internal List<XElement> Styles { get; set; }

        /// <summary>
        /// Starting index for this paragraph
        /// </summary>
        internal int StartIndex { get; }

        /// <summary>
        /// End index for this paragraph
        /// </summary>
        internal int EndIndex { get; }

        /// <summary>
        /// Constructor for the paragraph
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">XML for the paragraph</param>
        /// <param name="startIndex">Starting position in the doc</param>
        /// <param name="parentContainerType">Container parent type</param>
        internal Paragraph(DocX document, XElement xml, int startIndex, ContainerType parentContainerType = ContainerType.None)
            : base(document, xml)
        {
            ParentContainerType = parentContainerType;
            StartIndex = startIndex;
            EndIndex = startIndex + GetElementTextLength(Xml);
            Styles = new List<XElement>();
            Runs = Xml.Elements(DocxNamespace.Main + "r").ToList();
        }

        /// <summary>
        /// Gets or set this Paragraphs text alignment.
        /// </summary>
        public Alignment Alignment
        {
            get
            {
                var jc = ParaElement().Element(DocxNamespace.Main + "jc");
                if (jc != null && jc.TryGetEnumValue(out Alignment result))
                {
                    return result;
                }
                return Alignment.Left;
            }

            set
            {
                var pPr = ParaElement();
                var jc = pPr.Element(DocxNamespace.Main + "jc");

                if (value != Alignment.Left)
                {
                    if (jc == null)
                    {
                        pPr.Add(new XElement(DocxNamespace.Main + "jc",
                                    new XAttribute(DocxNamespace.Main + "val", value.GetEnumName())));
                    }
                    else
                    {
                        jc.SetAttributeValue(DocxNamespace.Main + "val", value.GetEnumName());
                    }
                }
                else
                {
                    jc?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or Sets the Direction of content in this Paragraph.
        /// </summary>
        public Direction Direction
        {
            get => ParaElement().Element(DocxNamespace.Main + "bidi") == null ? Direction.LeftToRight : Direction.RightToLeft;

            set
            {
                var pPr = ParaElement();
                var bidi = pPr.Element(DocxNamespace.Main + "bidi");

                if (value == Direction.RightToLeft)
                {
                    if (bidi == null)
                    {
                        pPr.Add(new XElement(DocxNamespace.Main + "bidi"));
                    }
                }
                else
                {
                    bidi?.Remove();
                }
            }
        }

        /// <summary>
        /// Returns a list of DocProperty elements in this document.
        /// </summary>
        public ReadOnlyCollection<DocProperty> DocumentProperties =>
            Xml.Descendants(DocxNamespace.Main + "fldSimple")
                .Select(el => new DocProperty(Document, el))
                .ToList().AsReadOnly();

        ///<summary>
        /// Returns table following the paragraph. Null if the following element isn't table.
        ///</summary>
        public Table FollowingTable { get; internal set; }

        /// <summary>
        /// Add a heading
        /// </summary>
        /// <param name="headingType"></param>
        public Paragraph Heading(HeadingType headingType)
        {
            StyleName = headingType.GetEnumName();
            return this;
        }

        /// <summary>
        /// Set the left indentation in cm for this Paragraph.
        /// </summary>
        public double IndentationLeft
        {
            get
            {
                var ind = ParaIndentationElement();
                var left = ind.Attribute(DocxNamespace.Main + "left");
                return left != null ? double.Parse(left.Value) / 20.0 : 0;
            }

            set
            {
                var ind = ParaIndentationElement();
                ind.SetAttributeValue(DocxNamespace.Main + "left", value * 20.0);
            }
        }

        /// <summary>
        /// Set the right indentation in cm for this Paragraph.
        /// </summary>
        public double IndentationRight
        {
            get
            {
                var ind = ParaIndentationElement();
                var right = ind.Attribute(DocxNamespace.Main + "right");
                return right != null ? double.Parse(right.Value)/20.0 : 0;
            }

            set
            {
                var ind = ParaIndentationElement();
                ind.SetAttributeValue(DocxNamespace.Main + "right", value * 20.0);
            }
        }

        /// <summary>
        /// Get or set the indentation of the first line of this Paragraph.
        /// </summary>
        public double IndentationFirstLine
        {
            get
            {
                var ind = ParaIndentationElement();
                var firstLine = ind.Attribute(DocxNamespace.Main + "firstLine");
                return firstLine != null ? double.Parse(firstLine.Value)/20.0 : 0;
            }

            set
            {
                var ind = ParaIndentationElement();

                // Remove any hanging indentation and set the firstLine indent.
                ind.Attribute(DocxNamespace.Main + "hanging")?.Remove();
                ind.SetAttributeValue(DocxNamespace.Main + "firstLine", value * 20f);
            }
        }

        /// <summary>
        /// Get or set the indentation of all but the first line of this Paragraph.
        /// </summary>
        public double IndentationHanging
        {
            get
            {
                var ind = ParaIndentationElement();
                var hanging = ind.Attribute(DocxNamespace.Main + "hanging");
                return hanging != null ? double.Parse(hanging.Value) / 20.0 : 0;
            }

            set
            {
                var ind = ParaIndentationElement();

                // Remove any firstLine indent and set hanging.
                ind.Attribute(DocxNamespace.Main + "firstLine")?.Remove();
                ind.SetAttributeValue(DocxNamespace.Main + "hanging", value * 20);
            }
        }

        /// <summary>
        /// If this element is a list item, get the indentation level of the list item.
        /// </summary>
        public int? IndentLevel
        {
            get
            {
                if (IsListItem)
                {
                    string value = ParagraphNumberProperties.FirstLocalNameDescendant("ilvl").GetVal();
                    if (value != null && int.TryParse(value, out int result))
                        return result;
                }

                return null;
            }
        }

        /// <summary>
        /// True to keep with the next element on the page.
        /// </summary>
        public bool ShouldKeepWithNext => ParaElement().Element(DocxNamespace.Main + "keepNext") != null;

        /// <summary>
        /// Determine if this paragraph is a list element.
        /// </summary>
        public bool IsListItem => ParagraphNumberProperties != null;

        /// <summary>
        /// Get or set the paragraph line spacing
        /// </summary>
        public double LineSpacing
        {
            get
            {
                var spacing = ParaElement().Element(DocxNamespace.Main + "spacing");
                var line = spacing?.Attribute(DocxNamespace.Main + "line");
                return line != null && double.TryParse(line.Value, out double value) ? value / 20.0 : 12;
            }

            set => Spacing(value);
        }

        /// <summary>
        /// Get or set the line spacing before this paragraph
        /// </summary>
        public double LineSpacingAfter
        {
            get
            {
                var spacing = ParaElement().Element(DocxNamespace.Main + "spacing");
                var line = spacing?.Attribute(DocxNamespace.Main + "after");
                return line != null && double.TryParse(line.Value, out double value) ? value / 20.0 : 0;
            }

            set => SpacingAfter(value);
        }

        /// <summary>
        /// Get or set the line spacing before this paragraph
        /// </summary>
        public double LineSpacingBefore
        {
            get
            {
                var spacing = ParaElement().Element(DocxNamespace.Main + "spacing");
                var line = spacing?.Attribute(DocxNamespace.Main + "before");
                return line != null && double.TryParse(line.Value, out double value) ? value / 20.0 : 0;
            }

            set => SpacingBefore(value);
        }

        /// <summary>
        /// Return the type of list contained in this paragraph.
        /// </summary>
        public ListItemType ListItemType => GetListItemType();

        /// <summary>
        /// Fetch the paragraph number properties for a list element.
        /// </summary>
        public XElement ParagraphNumberProperties
        {
            get
            {
                var node = Xml.FirstLocalNameDescendant("numPr");
                var numIdAttr = node?.FirstLocalNameDescendant("numId");
                return numIdAttr != null && int.TryParse(numIdAttr.GetVal(), out int result) && result > 0
                    ? node
                    : null;
            }
        }

        /// <summary>
        /// Get the ListItemType property for the paragraph.
        /// Defaults to numbered if a list is found but the type is not specified
        /// </summary>
        private ListItemType GetListItemType()
        {
            var numProperties = ParagraphNumberProperties;
            if (numProperties == null)
                return ListItemType.None;

            string level = numProperties.Element(DocxNamespace.Main + "ilvl").GetVal();
            string numIdRef = numProperties.Element(DocxNamespace.Main + "numId").GetVal();

            // Find the number definition instance. We map <w:num> to <w:abstractNum>
            var numNode = Document.numberingDoc.LocalNameDescendants("num")?.FindByAttrVal(DocxNamespace.Main + "numId", numIdRef);
            if (numNode == null)
            {
                throw new Exception(
                    $"Number reference w:numId('{numIdRef}') used in document but not defined in numbering.xml");
            }

            // Get the abstractNumId
            string absNumId = numNode.FirstLocalNameDescendant("abstractNumId").GetVal();

            // Find the numbering style section
            var absNumNode = Document.numberingDoc.LocalNameDescendants("abstractNum")
                .FindByAttrVal(DocxNamespace.Main + "abstractNumId", absNumId);

            // Get the numbering format.
            var format = absNumNode.LocalNameDescendants("lvl")
                .FindByAttrVal(DocxNamespace.Main + "ilvl", level)
                .FirstLocalNameDescendant("numFmt");

            return format.TryGetEnumValue(out ListItemType result)
                ? result
                : ListItemType.Numbered;
        }

        /// <summary>
        /// Parent container type
        /// </summary>
        public ContainerType ParentContainerType { get; set; }

        /// <summary>
        /// Returns a list of all Pictures in a Paragraph.
        /// </summary>
        public List<Picture> Pictures => (
                    from p in Xml.LocalNameDescendants("drawing")
                    let id = p.FirstLocalNameDescendant("blip").AttributeValue(DocxNamespace.RelatedDoc + "embed")
                    where id != null
                    let img = new Image(Document, PackagePart.GetRelationship(id))
                    select new Picture(Document, p, img)
                ).Union(
                    from p in Xml.LocalNameDescendants("pict")
                    let id = p.FirstLocalNameDescendant("imagedata").AttributeValue(DocxNamespace.RelatedDoc + "id")
                    where id != null
                    let img = new Image(Document, PackagePart.GetRelationship(id))
                    select new Picture(Document, p, img)
                ).ToList();

        ///<summary>
        /// The style name of the paragraph.
        ///</summary>
        public string StyleName
        {
            get
            {
                var styleElement = ParaElement().Element(DocxNamespace.Main + "pStyle");
                var attr = styleElement?.Attribute(DocxNamespace.Main + "val");
                return attr != null && !string.IsNullOrEmpty(attr.Value) ? attr.Value : "Normal";
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                    value = "Normal";

                var pPr = ParaElement();
                var styleElement = pPr.GetOrCreateElement(DocxNamespace.Main + "pStyle");
                styleElement.SetAttributeValue(DocxNamespace.Main + "val", value);
            }
        }

        /// <summary>
        /// Gets the text value of this Paragraph.
        /// </summary>
        public string Text => HelperFunctions.GetText(Xml);

        /// <summary>
        /// Fluent method to set alignment
        /// </summary>
        /// <param name="alignment">Desired alignment</param>
        public Paragraph Align(Alignment alignment)
        {
            Alignment = alignment;
            return this;
        }

        /// <summary>
        /// Append text to this Paragraph.
        /// </summary>
        /// <param name="text">The text to append.</param>
        /// <returns>This Paragraph with the new text appended.</returns>
        public Paragraph Append(string text)
        {
            var newRuns = HelperFunctions.FormatInput(text, null);
            Xml.Add(newRuns);
            Runs = Xml.Elements(DocxNamespace.Main + "r").Reverse().Take(newRuns.Count).ToList();

            return this;
        }

        /// <summary>
        /// Appends a new bookmark to the paragraph
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <returns>This paragraph</returns>
        public Paragraph AppendBookmark(string bookmarkName)
        {
            Xml.Add(new XElement(DocxNamespace.Main + "bookmarkStart",
                        new XAttribute(DocxNamespace.Main + "id", 0),
                        new XAttribute(DocxNamespace.Main + "name", bookmarkName)));

            Xml.Add(new XElement(DocxNamespace.Main + "bookmarkEnd",
                        new XAttribute(DocxNamespace.Main + "id", 0),
                        new XAttribute(DocxNamespace.Main + "name", bookmarkName)));

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
            var oMathPara = new XElement(DocxNamespace.Math + "oMathPara",
                    new XElement(DocxNamespace.Math + "oMath",
                        new XElement(DocxNamespace.Main + "r",
                            new Formatting { FontFamily = new FontFamily("Cambria Math") }.Xml,
                            new XElement(DocxNamespace.Math + "t", equation)
                        )
                    )
                );

            // Add equation element into paragraph xml and update runs collection
            Xml.Add(oMathPara);
            Runs = Xml.Elements(DocxNamespace.Math + "oMathPara").ToList();

            // Return paragraph with equation
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
            _ = HelperFunctions.EnsureRelsPathExists(this);

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
                Run run = GetFirstRunEffectedByEdit(charIndex);
                if (run == null)
                {
                    // Add this hyperlink as the last element.
                    Xml.Add(hyperlink.Xml);
                }
                else
                {
                    // Split this run at the point you want to insert
                    var splitRun = Run.SplitRun(run, charIndex);

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
            // Ensure the owner document has the hyperlink styles.
            Document.AddHyperlinkStyleIfNotPresent();

            // Set the document/package for the hyperlink to be owned by the document
            hyperlink.Document = Document;
            hyperlink.PackagePart = Document.PackagePart;

            // Check to see if the rels file exists and create it if not.
            _ = HelperFunctions.EnsureRelsPathExists(this);

            // Check to see if a rel for this Hyperlink exists, create it if not.
            _ = hyperlink.GetOrCreateRelationship();
            Xml.Add(hyperlink.Xml);
            Runs = Xml.Elements().Last().Elements(DocxNamespace.Main + "r").ToList();

            return this;
        }

        /// <summary>
        /// Append a new line to this Paragraph.
        /// </summary>
        /// <returns>This Paragraph with a new line appended.</returns>
        public Paragraph AppendLine()
        {
            return Append("\n");
        }

        /// <summary>
        /// Append text on a new line to this Paragraph.
        /// </summary>
        /// <param name="text">The text to append.</param>
        /// <returns>This Paragraph with the new text appended.</returns>
        public Paragraph AppendLine(string text)
        {
            return Append("\n" + text);
        }

        /// <summary>
        /// Append a PageCount place holder onto the end of a Paragraph.
        /// </summary>
        /// <param name="pageNumberFormat">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        public void AppendPageCount(PageNumberFormat pageNumberFormat)
        {
            var fldSimple = new XElement(DocxNamespace.Main + "fldSimple");

            fldSimple.Add(pageNumberFormat == PageNumberFormat.Normal
                ? new XAttribute(DocxNamespace.Main + "instr", @" NUMPAGES   \* MERGEFORMAT ")
                : new XAttribute(DocxNamespace.Main + "instr", @" NUMPAGES  \* ROMAN  \* MERGEFORMAT "));

            var rsid = Document.RevisionId;

            var content = XElement.Parse(
             $@"<w:r w:rsidR='{rsid}' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof />
                   </w:rPr>
                   <w:t>1</w:t>
               </w:r>"
            );

            fldSimple.Add(content);
            Xml.Add(fldSimple);
        }

        /// <summary>
        /// Append a PageNumber place holder onto the end of a Paragraph.
        /// </summary>
        /// <param name="pageNumberFormat">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        public void AppendPageNumber(PageNumberFormat pageNumberFormat)
        {
            var fldSimple = new XElement(DocxNamespace.Main + "fldSimple");

            fldSimple.Add(pageNumberFormat == PageNumberFormat.Normal
                ? new XAttribute(DocxNamespace.Main + "instr", @" PAGE   \* MERGEFORMAT ")
                : new XAttribute(DocxNamespace.Main + "instr", @" PAGE  \* ROMAN  \* MERGEFORMAT "));

            var rsid = Document.RevisionId;

            var content = XElement.Parse(
             $@"<w:r w:rsidR='{rsid}' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof />
                   </w:rPr>
                   <w:t>1</w:t>
               </w:r>"
            );

            fldSimple.Add(content);
            Xml.Add(fldSimple);
        }

        /// <summary>
        /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
        /// </summary>
        /// <param name="p">The Picture to append.</param>
        /// <returns>The Paragraph with the Picture now appended.</returns>
        public Paragraph AppendPicture(Picture p)
        {
            // Check to see if the rels file exists and create it if not.
            _ = HelperFunctions.EnsureRelsPathExists(this);

            // Check to see if a rel for this Picture exists, create it if not.
            string id = GetOrCreateRelationship(p);

            // Add the Picture Xml to the end of the paragraph Xml.
            Xml.Add(p.Xml);

            // Extract the attribute id from the Pictures Xml.
            var attributeId = Xml.Elements().Last().Descendants()
                            .Where(e => e.Name.LocalName.Equals("blip"))
                            .Select(e => e.Attribute(DocxNamespace.RelatedDoc + "embed"))
                            .Single();

            // Set its value to the Pictures relationships id.
            attributeId.SetValue(id);

            // For formatting such as .Bold()
            Runs = Xml.Elements(DocxNamespace.Main + "r").Reverse()
                      .Take(p.Xml.Elements(DocxNamespace.Main + "r")
                      .Count()).ToList();

            return this;
        }

        /// <summary>
        /// Append text to this Paragraph and then make it bold.
        /// </summary>
        /// <returns>This Paragraph with the last appended text bold.</returns>
        public Paragraph Bold()
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "b", string.Empty, null);
            return this;
        }

        /// <summary>
        /// Append text to this Paragraph and then set it to full caps.
        /// </summary>
        /// <param name="capsStyle">The caps style to apply to the last appended text.</param>
        /// <returns>This Paragraph with the last appended text's caps style changed.</returns>
        public Paragraph CapsStyle(CapsStyle capsStyle)
        {
            if (capsStyle != DXPlus.CapsStyle.None)
            {
                ApplyTextFormattingProperty(DocxNamespace.Main + capsStyle.GetEnumName(), string.Empty, null);
            }

            return this;
        }

        /// <summary>
        /// Append text to this Paragraph and then color it.
        /// </summary>
        /// <param name="color">A color to use on the appended text.</param>
        /// <returns>This Paragraph with the last appended text colored.</returns>
        public Paragraph Color(Color color)
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "color", string.Empty, 
                new XAttribute(DocxNamespace.Main + "val", color.ToHex()));
            return this;
        }

        /// <summary>
        /// Set the culture of the preceding text
        /// </summary>
        /// <param name="culture">The CultureInfo for text</param>
        /// <returns>This Paragraph in current culture</returns>
        public Paragraph Culture(CultureInfo culture)
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "lang", string.Empty,
                new XAttribute(DocxNamespace.Main + "val", culture.Name));
            return this;
        }

        /// <summary>
        /// Set the culture of the preceding text
        /// </summary>
        /// <returns>This Paragraph in current culture</returns>
        public Paragraph Culture() => Culture(CultureInfo.CurrentCulture);

        /// <summary>
        /// Find all instances of a string in this paragraph and return their indexes in a List.
        /// </summary>
        /// <param name="text">The string to find</param>
        /// <param name="ignoreCase">True to ignore case in the search</param>
        /// <returns>A list of indexes.</returns>
        public IEnumerable<int> FindAll(string text, bool ignoreCase)
        {
            var mc = Regex.Matches(Text, Regex.Escape(text), ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
            return mc.Select(m => m.Index);
        }

        /// <summary>
        ///  Find all unique instances of the given Regex Pattern
        /// </summary>
        /// <param name="regex">Regex to match</param>
        /// <returns>Index and matched text</returns>
        public IEnumerable<(int index, string text)> FindPattern(Regex regex)
        {
            var mc = regex.Matches(Text);
            return mc.Select(m => (index: m.Index, text: m.Value));
        }

        /// <summary>
        /// Set the font for the preceding text.
        /// </summary>
        /// <param name="fontFamily">The font to use for the appended text.</param>
        /// <returns>This Paragraph with the last appended text's font changed.</returns>
        public Paragraph Font(FontFamily fontFamily)
        {
            ApplyTextFormattingProperty(
                DocxNamespace.Main + "rFonts",
                string.Empty,
                new[] {
                    new XAttribute(DocxNamespace.Main + "ascii", fontFamily.Name),
                    new XAttribute(DocxNamespace.Main + "hAnsi", fontFamily.Name),
                    new XAttribute(DocxNamespace.Main + "cs", fontFamily.Name)
                }
            );

            return this;
        }

        /// <summary>
        /// Set the font size for the appended text in this paragraph
        /// </summary>
        /// <param name="fontSize">The font size to use for the appended text.</param>
        /// <returns>Paragraph with the last appended text resized.</returns>
        public Paragraph FontSize(double fontSize)
        {
            // [0-1638] rounded to nearest half.
            fontSize = Math.Min(Math.Max(0, fontSize), 1638.0);
            fontSize = Math.Round(fontSize * 2, MidpointRounding.AwayFromZero) / 2;

            ApplyTextFormattingProperty(DocxNamespace.Main + "sz", string.Empty, new XAttribute(DocxNamespace.Main + "val", fontSize * 2));
            ApplyTextFormattingProperty(DocxNamespace.Main + "szCs", string.Empty, new XAttribute(DocxNamespace.Main + "val", fontSize * 2));

            return this;
        }

        /// <summary>
        /// Retrieve all the bookmarks in this paragraph
        /// </summary>
        /// <returns>Enumerable of bookmark objects</returns>
        public IEnumerable<Bookmark> GetBookmarks()
        {
            return Xml.Descendants(DocxNamespace.Main + "bookmarkStart")
                        .Select(x => x.Attribute(DocxNamespace.Main + "name"))
                        .Select(x => new Bookmark(x.Value, this));
        }

        /// <summary>
        /// Append text to this Paragraph and then hide it.
        /// </summary>
        /// <returns>This Paragraph with the last appended text hidden.</returns>
        public Paragraph Hide()
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "vanish", string.Empty, null);
            return this;
        }

        /// <summary>
        /// Highlights the given paragraph/line
        /// </summary>
        ///<param name="highlight">The highlight to apply to the last appended text.</param>
        /// <returns>This Paragraph with the last appended text highlighted.</returns>
        public Paragraph Highlight(Highlight highlight)
        {
            if (highlight != DXPlus.Highlight.None)
            {
                ApplyTextFormattingProperty(DocxNamespace.Main + "highlight", string.Empty,
                    new XAttribute(DocxNamespace.Main + "val", highlight.GetEnumName()));
            }

            return this;
        }

        /// <summary>
        /// Insert a text block at a bookmark
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <param name="toInsert">Text to insert</param>
        public bool InsertAtBookmark(string bookmarkName, string toInsert)
        {
            var bookmark = Xml.Descendants(DocxNamespace.Main + "bookmarkStart")
                    .SingleOrDefault(x => x.AttributeValue(DocxNamespace.Main + "name") == bookmarkName);
            if (bookmark != null)
            {
                var run = HelperFunctions.FormatInput(toInsert, null);
                bookmark.AddBeforeSelf(run);
                HelperFunctions.RenumberIds(Document);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
        /// </summary>
        /// <param name="cp">The custom property to display.</param>
        /// <param name="trackChanges"></param>
        /// <param name="formatting">The formatting to use for this text.</param>
        public DocProperty AddDocumentProperty(CustomProperty cp, bool trackChanges = false, Formatting formatting = null)
        {
            var formattingXml = formatting?.Xml;

            var e = new XElement(DocxNamespace.Main + "fldSimple",
                new XAttribute(DocxNamespace.Main + "instr", $@"DOCPROPERTY {cp.Name} \* MERGEFORMAT"),
                    new XElement(DocxNamespace.Main + "r",
                        new XElement(DocxNamespace.Main + "t", formattingXml, cp.Value))
            );

            var xml = e;
            if (trackChanges)
            {
                DateTime now = DateTime.Now.ToUniversalTime();
                e = HelperFunctions.CreateEdit(EditType.Ins, new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc), e);
            }

            Xml.Add(e);

            return new DocProperty(Document, xml);
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
            var fldSimple = new XElement(DocxNamespace.Main + "fldSimple");

            fldSimple.Add(pnf == PageNumberFormat.Normal
                ? new XAttribute(DocxNamespace.Main + "instr", @" NUMPAGES   \* MERGEFORMAT ")
                : new XAttribute(DocxNamespace.Main + "instr", @" NUMPAGES  \* ROMAN  \* MERGEFORMAT "));

            var rsid = Document.RevisionId;
            var content = XElement.Parse(
             $@"<w:r w:rsidR='{rsid}' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof />
                   </w:rPr>
                   <w:t>1</w:t>
               </w:r>"
            );

            fldSimple.Add(content);

            if (index == 0)
            {
                Xml.AddFirst(fldSimple);
            }
            else
            {
                Run r = GetFirstRunEffectedByEdit(index);
                var splitEdit = SplitEdit(r.Xml, index, EditType.Ins);
                r.Xml.ReplaceWith(splitEdit[0], fldSimple, splitEdit[1]);
            }
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
            var fldSimple = new XElement(DocxNamespace.Main + "fldSimple");

            fldSimple.Add(pnf == PageNumberFormat.Normal
                ? new XAttribute(DocxNamespace.Main + "instr", @" PAGE   \* MERGEFORMAT ")
                : new XAttribute(DocxNamespace.Main + "instr", @" PAGE  \* ROMAN  \* MERGEFORMAT "));

            var rsid = Document.RevisionId;
            var content = XElement.Parse(
             $@"<w:r w:rsidR='{rsid}' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                   <w:rPr>
                       <w:noProof />
                   </w:rPr>
                   <w:t>1</w:t>
               </w:r>"
            );

            fldSimple.Add(content);

            if (index == 0)
            {
                Xml.AddFirst(fldSimple);
            }
            else
            {
                Run r = GetFirstRunEffectedByEdit(index);
                var splitEdit = SplitEdit(r.Xml, index, EditType.Ins);
                r.Xml.ReplaceWith(splitEdit[0], fldSimple, splitEdit[1]);
            }
        }

        /// <summary>
        /// Insert a Picture into a Paragraph at the given text index.
        /// If not index is provided defaults to 0.
        /// </summary>
        /// <param name="picture">The Picture to insert.</param>
        /// <param name="index">The text index to insert at.</param>
        /// <returns>The modified Paragraph.</returns>
        public Paragraph InsertPicture(Picture picture, int index = 0)
        {
            // Check to see if the rels file exists and create it if not.
            _ = HelperFunctions.EnsureRelsPathExists(this);

            // Check to see if a rel for this Picture exists, create it if not.
            string id = GetOrCreateRelationship(picture);

            XElement pictureElement;
            if (index == 0)
            {
                // Add this hyperlink as the last element.
                Xml.AddFirst(picture.Xml);

                // Extract the picture back out of the DOM.
                pictureElement = (XElement)Xml.FirstNode;
            }
            else
            {
                // Get the first run effected by this Insert
                Run run = GetFirstRunEffectedByEdit(index);
                if (run == null)
                {
                    // Add this picture as the last element.
                    Xml.Add(picture.Xml);

                    // Extract the picture back out of the DOM.
                    pictureElement = (XElement)Xml.LastNode;
                }
                else
                {
                    // Split this run at the point you want to insert
                    XElement[] splitRun = Run.SplitRun(run, index);

                    // Replace the original run.
                    run.Xml.ReplaceWith(splitRun[0], picture.Xml, splitRun[1]);

                    // Get the first run effected by this Insert
                    run = GetFirstRunEffectedByEdit(index);

                    // The picture has to be the next element, extract it back out of the DOM.
                    pictureElement = (XElement)run.Xml.NextNode;
                }
            }
            // Extract the attribute id from the Pictures Xml.
            XAttribute attributeId = pictureElement.LocalNameDescendants("blip")
                    .Select(e => e.Attribute(DocxNamespace.RelatedDoc + "embed"))
                    .Single();

            // Set its value to the Pictures relationships id.
            attributeId.SetValue(id);

            return this;
        }

        /// <summary>
        /// Inserts a string into a Paragraph with the specified formatting.
        /// </summary>
        public void InsertText(string value, Formatting formatting = null)
        {
            if (formatting == null)
                formatting = new Formatting();

            Xml.Add(HelperFunctions.FormatInput(value, formatting.Xml));
            HelperFunctions.RenumberIds(Document);
        }

        /// <summary>
        /// Inserts a string into a Paragraph with the specified formatting at the given index.
        /// </summary>
        /// <param name="index">The index position of the insertion.</param>
        /// <param name="value">The System.String to insert.</param>
        /// <param name="trackChanges">Flag this insert as a change.</param>
        /// <param name="formatting">The text formatting.</param>
        public void InsertText(int index, string value, bool trackChanges = false, Formatting formatting = null)
        {
            // Timestamp to mark the start of insert
            DateTime now = DateTime.Now.ToUniversalTime();
            DateTime insertDate = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

            // Get the first run effected by this Insert
            var run = GetFirstRunEffectedByEdit(index);
            if (run == null)
            {
                object insert = HelperFunctions.FormatInput(value, formatting?.Xml);

                if (trackChanges)
                    insert = HelperFunctions.CreateEdit(EditType.Ins, insertDate, insert);

                Xml.Add(insert);
            }
            else
            {
                object newRuns;
                var rPr = run.Xml.GetRunProps();
                if (formatting != null)
                {
                    Formatting oldFormat = null;
                    if (rPr != null)
                    {
                        oldFormat = Formatting.Parse(rPr);
                    }

                    // Create the formatting - it's a mix of the original formatting + the passed formatting info.
                    Formatting newFormat;
                    if (oldFormat != null)
                    {
                        newFormat = oldFormat.Clone();

                        if (formatting.Bold.HasValue)
                            newFormat.Bold = formatting.Bold;
                        if (formatting.CapsStyle.HasValue)
                            newFormat.CapsStyle = formatting.CapsStyle;
                        if (formatting.FontColor.HasValue)
                            newFormat.FontColor = formatting.FontColor;
                        newFormat.FontFamily = formatting.FontFamily;
                        if (formatting.Hidden.HasValue)
                            newFormat.Hidden = formatting.Hidden;
                        if (formatting.Highlight.HasValue)
                            newFormat.Highlight = formatting.Highlight;
                        if (formatting.Italic.HasValue)
                            newFormat.Italic = formatting.Italic;
                        if (formatting.Kerning.HasValue)
                            newFormat.Kerning = formatting.Kerning;
                        newFormat.Language = formatting.Language;
                        if (formatting.Misc.HasValue)
                            newFormat.Misc = formatting.Misc;
                        if (formatting.PercentageScale.HasValue)
                            newFormat.PercentageScale = formatting.PercentageScale;
                        if (formatting.Position.HasValue)
                            newFormat.Position = formatting.Position;
                        if (formatting.Script.HasValue)
                            newFormat.Script = formatting.Script;
                        if (formatting.Size.HasValue)
                            newFormat.Size = formatting.Size;
                        if (formatting.Spacing.HasValue)
                            newFormat.Spacing = formatting.Spacing;
                        if (formatting.Strikethrough.HasValue)
                            newFormat.Strikethrough = formatting.Strikethrough;
                        if (formatting.UnderlineColor.HasValue)
                            newFormat.UnderlineColor = formatting.UnderlineColor;
                        if (formatting.UnderlineStyle.HasValue)
                            newFormat.UnderlineStyle = formatting.UnderlineStyle;
                    }
                    else
                    {
                        newFormat = formatting;
                    }

                    newRuns = HelperFunctions.FormatInput(value, newFormat.Xml);
                }
                else
                {
                    newRuns = HelperFunctions.FormatInput(value, rPr);
                }

                // The parent of this Run
                var parentElement = run.Xml.Parent;
                object insert = newRuns;

                switch (parentElement.Name.LocalName)
                {
                    case "ins":
                        // The datetime that this ins was created
                        var parentInsertDate = DateTime.Parse(parentElement.Attribute(DocxNamespace.Main + "date").Value);
                        // Special case: You want to track changes, and the first Run effected by this insert
                        // has a datetime stamp equal to now.
                        if (trackChanges && parentInsertDate.CompareTo(insertDate) == 0)
                            goto default;

                        goto case "del";

                    case "del":
                        if (trackChanges)
                            insert = HelperFunctions.CreateEdit(EditType.Ins, insertDate, newRuns);

                        // Split this Edit at the point you want to insert
                        var splitEdit = SplitEdit(parentElement, index, EditType.Ins);

                        // Replace the original run
                        parentElement.ReplaceWith(splitEdit[0], insert, splitEdit[1]);
                        break;

                    default:
                        if (trackChanges && !parentElement.Name.LocalName.Equals("ins"))
                        {
                            _ = HelperFunctions.CreateEdit(EditType.Ins, insertDate, newRuns);
                        }
                        else
                        {
                            // Split this run at the point you want to insert
                            var splitRun = Run.SplitRun(run, index);
                            run.Xml.ReplaceWith(splitRun[0], insert, splitRun[1]);
                        }
                        break;
                }
            }

            HelperFunctions.RenumberIds(Document);
        }

        /// <summary>
        /// Append text to this Paragraph and then make it italic.
        /// </summary>
        /// <returns>This Paragraph with the last appended text italic.</returns>
        public Paragraph Italic()
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "i", string.Empty, null);
            return this;
        }

        /// <summary>
        /// Keep all lines in this paragraph together on a page
        /// </summary>
        public Paragraph KeepLinesTogether(bool keepTogether = true)
        {
            var pPr = ParaElement();
            var keepLinesElement = pPr.Element(DocxNamespace.Main + "keepLines");
            if (keepLinesElement == null && keepTogether)
            {
                pPr.Add(new XElement(DocxNamespace.Main + "keepLines"));
            }
            else if (!keepTogether)
            {
                keepLinesElement?.Remove();
            }
            return this;
        }

        /// <summary>
        /// This paragraph will be kept on the same page as the next paragraph
        /// </summary>
        /// <param name="keepWithNext"></param>
        public Paragraph KeepWithNext(bool keepWithNext = true)
        {
            var pPr = ParaElement();
            var keepWithNextElement = pPr.Element(DocxNamespace.Main + "keepNext");
            if (keepWithNextElement == null && keepWithNext)
            {
                pPr.Add(new XElement(DocxNamespace.Main + "keepNext"));
            }
            else if (!keepWithNext)
            {
                keepWithNextElement?.Remove();
            }
            return this;
        }

        /// <summary>
        /// Set the kerning value for the paragraph
        /// </summary>
        /// <param name="kerning"></param>
        /// <returns></returns>
        public Paragraph Kerning(int kerning)
        {
            int[] validValues = {8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72};

            if (!validValues.Contains(kerning))
                throw new ArgumentOutOfRangeException(nameof(kerning), "Value must be one of the following: " + string.Join(',', validValues.Select(i => i.ToString())));

            ApplyTextFormattingProperty(DocxNamespace.Main + "kern", string.Empty, new XAttribute(DocxNamespace.Main + "val", kerning * 2));
            return this;
        }

        /// <summary>
        /// Set one of the misc. properties
        /// </summary>
        /// <param name="misc">The miscellaneous property to set.</param>
        /// <returns>This Paragraph with the last appended text changed by a miscellaneous property.</returns>
        public Paragraph Misc(Misc misc)
        {
            switch (misc)
            {
                case DXPlus.Misc.None:
                    break;

                case DXPlus.Misc.OutlineShadow:
                    ApplyTextFormattingProperty(DocxNamespace.Main + "outline", string.Empty, null);
                    ApplyTextFormattingProperty(DocxNamespace.Main + "shadow", string.Empty, null);
                    break;

                default:
                    ApplyTextFormattingProperty(DocxNamespace.Main + misc.GetEnumName(), string.Empty, null);
                    break;
            }

            return this;
        }

        /// <summary>
        /// Set the percentage scale for the paragraph
        /// </summary>
        /// <param name="percentageScale"></param>
        /// <returns>Paragraph</returns>
        public Paragraph PercentageScale(int percentageScale)
        {
            int[] validValues = {200, 150, 100, 90, 80, 66, 50, 33};

            if (!validValues.Contains(percentageScale))
                throw new ArgumentOutOfRangeException(nameof(percentageScale), "Value must be one of the following: " + string.Join(',', validValues.Select(i => i.ToString())));

            ApplyTextFormattingProperty(DocxNamespace.Main + "w", string.Empty, new XAttribute(DocxNamespace.Main + "val", percentageScale));

            return this;
        }

        /// <summary>
        /// Set the vertical position of the paragraph in half-points
        /// </summary>
        /// <param name="position"></param>
        /// <returns>Paragraph</returns>
        public Paragraph Position(double position)
        {
            position = Math.Min(Math.Max(-1585.0, position), 1585.0);
            ApplyTextFormattingProperty(DocxNamespace.Main + "position", string.Empty, new XAttribute(DocxNamespace.Main + "val", position * 2));

            return this;
        }

        /// <summary>
        /// Remove this Paragraph from the document.
        /// </summary>
        /// <param name="trackChanges">Should this remove be tracked as a change?</param>
        public void Remove(bool trackChanges)
        {
            if (trackChanges)
            {
                var now = DateTime.Now.ToUniversalTime();

                var elements = Xml.Elements().ToList();
                var temp = new List<XElement>();
                
                for (int i = 0; i < elements.Count; i++)
                {
                    var e = elements[i];
                    if (e.Name.LocalName != "del")
                    {
                        temp.Add(e);
                        e.Remove();
                    }
                    else if (temp.Count > 0)
                    {
                        e.AddBeforeSelf(HelperFunctions.CreateEdit(EditType.Del, now, temp.Elements()));
                        temp.Clear();
                    }
                }

                if (temp.Count > 0)
                {
                    Xml.Add(HelperFunctions.CreateEdit(EditType.Del, now, temp));
                }
            }
            else
            {
                // If this is the only Paragraph in the Cell then we cannot remove it.
                if (Xml.Parent?.Name.LocalName == "tc" && Xml.Parent.Elements(DocxNamespace.Main + "p").Count() == 1)
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
                throw new ArgumentOutOfRangeException(nameof(index));
        }

        /// <summary>
        /// Removes characters from a DXPlus.DocX.Paragraph.
        /// </summary>
        /// <param name="index">The position to begin deleting characters.</param>
        /// <param name="count">The number of characters to delete</param>
        /// <param name="trackChanges">Track changes</param>
        public void RemoveText(int index, int count, bool trackChanges = false)
        {
            // Timestamp to mark the start of insert
            DateTime now = DateTime.Now.ToUniversalTime();
            DateTime removeDate = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

            // The number of characters processed so far
            int processed = 0;

            do
            {
                // Get the first run effected by this Remove
                var run = GetFirstRunEffectedByEdit(index, EditType.Del);

                // The parent of this Run
                var parentElement = run.Xml.Parent;
                switch (parentElement?.Name.LocalName)
                {
                    case "ins":
                        {
                            var splitEditBefore = SplitEdit(parentElement, index, EditType.Del);
                            int min = Math.Min(count - processed, run.Xml.ElementsAfterSelf().Sum(GetElementTextLength));
                            var splitEditAfter = SplitEdit(parentElement, index + min, EditType.Del);
                            var temp = SplitEdit(splitEditBefore[1], index + min, EditType.Del)[0];
                            object middle = HelperFunctions.CreateEdit(EditType.Del, removeDate, temp.Elements());
                            processed += GetElementTextLength((XElement) middle);

                            if (!trackChanges)
                            {
                                middle = null;
                            }

                            parentElement.ReplaceWith(splitEditBefore[0], middle, splitEditAfter[1]);
                            processed += GetElementTextLength((XElement) middle);
                        }
                        break;

                    case "del":
                        if (trackChanges)
                        {
                            // You cannot delete from a deletion, advance processed to the end of this del
                            processed += GetElementTextLength(parentElement);
                        }
                        else
                        {
                            goto case "ins";
                        }
                        break;

                    default:
                        {
                            var splitRunBefore = Run.SplitRun(run, index, EditType.Del);
                            int min = Math.Min(index + (count - processed), run.EndIndex);
                            var splitRunAfter = Run.SplitRun(run, min, EditType.Del);

                            object middle = HelperFunctions.CreateEdit(EditType.Del, removeDate, new List<XElement> {
                                Run.SplitRun(new Run(Document, splitRunBefore[1], run.StartIndex + GetElementTextLength(splitRunBefore[0])), min, EditType.Del)[0]
                            });
                            processed += GetElementTextLength((XElement) middle);

                            if (!trackChanges)
                            {
                                middle = null;
                            }

                            run.Xml.ReplaceWith(splitRunBefore[0], middle, splitRunAfter[1]);
                        }
                        break;
                }

                // If after this remove the parent element is empty, remove it.
                if (GetElementTextLength(parentElement) == 0 
                    && parentElement?.Parent != null && parentElement.Parent.Name.LocalName != "tc")
                {
                    // Need to make sure there is no drawing element within the parent element.
                    // Picture elements contain no text length but they are still content.
                    if (!parentElement.Descendants(DocxNamespace.Main + "drawing").Any())
                    {
                        parentElement.Remove();
                    }
                }
            }
            while (processed < count);

            HelperFunctions.RenumberIds(Document);
        }

        /// <summary>
        /// Removes characters from a DXPlus.DocX.Paragraph.
        /// </summary>
        /// <param name="index">The position to begin deleting characters.</param>
        /// <param name="trackChanges">Track changes</param>
        public void RemoveText(int index, bool trackChanges = false)
        {
            RemoveText(index, Text.Length - index, trackChanges);
        }

        /// <summary>
        /// Replaces all occurrences of a specified System.String in this instance, with another specified System.String.
        /// </summary>
        /// <param name="newValue">A System.String to replace all occurrences of oldValue.</param>
        /// <param name="oldValue">A System.String to be replaced.</param>
        /// <param name="options">A bitwise OR combination of RegexOption enumeration options.</param>
        /// <param name="trackChanges">Track changes</param>
        /// <param name="newFormatting">The formatting to apply to the text being inserted.</param>
        /// <param name="matchFormatting">The formatting that the text must match in order to be replaced.</param>
        /// <param name="fo">How should formatting be matched?</param>
        /// <param name="escapeRegEx">True if the oldValue needs to be escaped, otherwise false. If it represents a valid RegEx pattern this should be false.</param>
        /// <param name="useRegExSubstitutions">True if RegEx-like replace should be performed, i.e. if newValue contains RegEx substitutions. Does not perform named-group substitutions (only numbered groups).</param>
        public void ReplaceText(string oldValue, string newValue, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch, bool escapeRegEx = true, bool useRegExSubstitutions = false)
        {
            string tText = Text;
            var mc = Regex.Matches(tText, escapeRegEx ? Regex.Escape(oldValue) : oldValue, options);

            // Loop through the matches in reverse order
            foreach (var m in mc.Reverse())
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
                        var run = GetFirstRunEffectedByEdit(m.Index + processed);

                        // Get this runs properties
                        var rPr = run.Xml.GetRunProps(false) ?? new Formatting().Xml;

                        // Make sure that every formatting element in f.xml is also in this run,
                        // if this is not true, then their formatting does not match.
                        if (!HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, rPr, fo))
                        {
                            formattingMatch = false;
                            break;
                        }

                        // We have processed some characters, so update the counter.
                        processed += run.Value.Length;
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
                                var g = m.Groups[k];
                                if (g.Value.Length == 0)
                                    continue;

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
                        InsertText(m.Index + m.Length, repl, trackChanges, newFormatting);
                    }

                    if (m.Length > 0)
                    {
                        RemoveText(m.Index, m.Length, trackChanges);
                    }
                }
            }
        }

        /// <summary>
        /// Find pattern regex must return a group match.
        /// </summary>
        /// <param name="findPattern">Regex pattern that must include one group match. ie (.*)</param>
        /// <param name="regexMatchHandler">A func that accepts the matching find grouping text and returns a replacement value</param>
        /// <param name="trackChanges"></param>
        /// <param name="options"></param>
        /// <param name="newFormatting"></param>
        /// <param name="matchFormatting"></param>
        /// <param name="fo"></param>
        public void ReplaceText(string findPattern, Func<string, string> regexMatchHandler, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions fo = MatchFormattingOptions.SubsetMatch)
        {
            var matchCollection = Regex.Matches(Text, findPattern, options);

            // Loop through the matches in reverse order
            foreach (var match in matchCollection.Reverse())
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
                        var run = GetFirstRunEffectedByEdit(match.Index + processed);

                        // Get this runs properties
                        var rPr = run.Xml.GetRunProps(false) ?? new Formatting().Xml;
                        if (!HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, rPr, fo))
                        {
                            formattingMatch = false;
                            break;
                        }

                        // We have processed some characters, so update the counter.
                        processed += run.Value.Length;
                    } while (processed < match.Length);
                }

                // If the formatting matches, do the replace.
                if (formattingMatch)
                {
                    string newValue = regexMatchHandler.Invoke(match.Groups[1].Value);
                    InsertText(match.Index + match.Value.Length, newValue, trackChanges, newFormatting);
                    RemoveText(match.Index, match.Value.Length, trackChanges);
                }
            }
        }

        /// <summary>
        /// Append text to this Paragraph and then set it to superscript.
        /// </summary>
        /// <param name="script">The script style to apply to the last appended text.</param>
        /// <returns>This Paragraph with the last appended text's script style changed.</returns>
        public Paragraph Script(Script script)
        {
            if (script != DXPlus.Script.None)
            {
                ApplyTextFormattingProperty(DocxNamespace.Main + "vertAlign", string.Empty,
                    new XAttribute(DocxNamespace.Main + "val", script.GetEnumName()));
            }

            return this;
        }

        /// <summary>
        /// Set the line spacing for this paragraph manually.
        /// </summary>
        /// <param name="spacingType">The type of spacing to be set, can be either Before, After or Line (Standard line spacing).</param>
        /// <param name="spacing">A float value of the amount of spacing. Equals the value that van be set in Word using the "Line and Paragraph spacing" button.</param>
        public void SetLineSpacing(LineSpacingType spacingType, double spacing)
        {
            var pPr = ParaElement();
            var spacingElement = pPr.GetOrCreateElement(DocxNamespace.Main + "spacing");
            spacingElement.SetAttributeValue(DocxNamespace.Main + spacingType.GetEnumName(), spacing * 20.0);
        }

        /// <summary>
        /// Set the line spacing for this paragraph using the Auto value.
        /// </summary>
        /// <param name="spacingType">The type of spacing to be set automatically. Using Auto will set both Before and After. None will remove any linespacing.</param>
        public void SetLineSpacing(LineSpacingTypeAuto spacingType)
        {
            var pPr = ParaElement();
            var spacing = pPr.GetOrCreateElement(DocxNamespace.Main + "spacing");

            if (spacingType.Equals(LineSpacingTypeAuto.None))
            {
                spacing?.Remove();
            }
            else
            {
                const int spacingValue = 14 * 20; // 14pt space
               
                if (spacingType == LineSpacingTypeAuto.Auto)
                {
                    spacing.SetAttributeValue(DocxNamespace.Main + "after", spacingValue);
                    spacing.SetAttributeValue(DocxNamespace.Main + "afterAutospacing", 1);
                }

                string spacingTypeAttribute = spacingType == LineSpacingTypeAuto.AutoAfter ? "after" : "before";
                string autoSpacingTypeAttribute = spacingTypeAttribute + "Autospacing";

                spacing.SetAttributeValue(DocxNamespace.Main + autoSpacingTypeAttribute, 1);
                spacing.SetAttributeValue(DocxNamespace.Main + spacingTypeAttribute, spacingValue);
            }
        }

        public Paragraph Spacing(double spacing)
        {
            spacing *= 20.0; // spacing is in 20ths of a pt
            spacing = Math.Round(Math.Min(Math.Max(0, spacing), 1584.0), 1);
            
            ApplyTextFormattingProperty(DocxNamespace.Main + "spacing", string.Empty, new XAttribute(DocxNamespace.Main + "line", spacing));

            return this;
        }

        public Paragraph SpacingAfter(double spacing)
        {
            spacing *= 20.0; // spacing is in 20ths of a pt
            spacing = Math.Round(Math.Min(Math.Max(0, spacing), 1584.0), 1);

            ApplyTextFormattingProperty(DocxNamespace.Main + "spacing", string.Empty, new XAttribute(DocxNamespace.Main + "after", spacing));

            return this;
        }

        public Paragraph SpacingBefore(double spacing)
        {
            spacing *= 20.0; // spacing is in 20ths of a pt
            spacing = Math.Round(Math.Min(Math.Max(0, spacing), 1584.0), 1);

            ApplyTextFormattingProperty(DocxNamespace.Main + "spacing", string.Empty, new XAttribute(DocxNamespace.Main + "before", spacing));

            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="strike">The strike through style to used on the last appended text.</param>
        /// <returns>This Paragraph with the last appended text striked.</returns>
        public Paragraph StrikeThrough(Strikethrough strike)
        {
            string value = strike.GetEnumName();
            ApplyTextFormattingProperty(DocxNamespace.Main + value, string.Empty, value);

            return this;
        }

        /// <summary>
        /// Fluent method to set StyleName property
        /// </summary>
        /// <param name="styleName">Stylename</param>
        public Paragraph Style(string styleName)
        {
            StyleName = styleName;
            return this;
        }

        /// <summary>
        /// Append text to this Paragraph and then underline it using a color.
        /// </summary>
        /// <param name="underlineColor">The underline color to use, if no underline is set, a single line will be used.</param>
        /// <returns>This Paragraph with the last appended text underlined in a color.</returns>
        public Paragraph UnderlineColor(Color underlineColor)
        {
            foreach (var run in Runs)
            {
                var rPr = run.GetRunProps();
                var u = rPr.Element(DocxNamespace.Main + "u");
                if (u == null)
                {
                    rPr.SetElementValue(DocxNamespace.Main + "u", string.Empty);
                    u = rPr.GetOrCreateElement(DocxNamespace.Main + "u");
                    u.SetAttributeValue(DocxNamespace.Main + "val", "single");
                }

                u.SetAttributeValue(DocxNamespace.Main + "color", underlineColor.ToHex());
            }

            return this;
        }

        /// <summary>
        /// Append text to this Paragraph and then underline it.
        /// </summary>
        /// <param name="underlineStyle">The underline style to use for the appended text.</param>
        /// <returns>This Paragraph with the last appended text underlined.</returns>
        public Paragraph UnderlineStyle(UnderlineStyle underlineStyle)
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "u", string.Empty, 
                new XAttribute(DocxNamespace.Main + "val", underlineStyle.GetEnumName()));
            return this;
        }

        /// <summary>
        /// Validate that a bookmark exists
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <returns></returns>
        public bool BookmarkExists(string bookmarkName)
        {
            return GetBookmarks().Any(b => b.Name.Equals(bookmarkName));
        }

        /// <summary>
        /// Create a new Picture.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="id">A unique id that identifies an Image embedded in this document.</param>
        /// <param name="name">The name of this Picture.</param>
        /// <param name="description">The description of this Picture.</param>
        internal static Picture CreatePicture(DocX document, string id, string name, string description)
        {
            var part = document.Package.GetPart(document.PackagePart.GetRelationship(id).TargetUri);

            int newDocPrId = 1;
            var existingIds = (
                from bookmarkId in document.Xml.Descendants(DocxNamespace.Main + "bookmarkStart") 
                select bookmarkId.Attributes().FirstOrDefault(x => x.Name.LocalName == "id") into idAtt 
                where idAtt != null 
                select idAtt.Value
            ).ToList();

            while (existingIds.Contains(newDocPrId.ToString()))
            {
                newDocPrId++;
            }

            int cx, cy;

            using (var partStream = part.GetStream())
            using (var img = System.Drawing.Image.FromStream(partStream))
            {
                cx = img.Width * 9526;
                cy = img.Height * 9526;
            }

            var xml = XElement.Parse(
                $@"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:drawing xmlns = ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                        <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                            <wp:extent cx=""{cx}"" cy=""{cy}"" />
                            <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0"" />
                            <wp:docPr id=""{newDocPrId.ToString()}"" name=""{name}"" descr=""{description}"" />
                            <wp:cNvGraphicFramePr>
                                <a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1"" />
                            </wp:cNvGraphicFramePr>
                            <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                                <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                    <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                        <pic:nvPicPr>
                                        <pic:cNvPr id=""0"" name=""{name}"" />
                                            <pic:cNvPicPr />
                                        </pic:nvPicPr>
                                        <pic:blipFill>
                                            <a:blip r:embed=""{id}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>
                                            <a:stretch>
                                                <a:fillRect />
                                            </a:stretch>
                                        </pic:blipFill>
                                        <pic:spPr>
                                            <a:xfrm>
                                                <a:off x=""0"" y=""0"" />
                                                <a:ext cx=""{cx}"" cy=""{cy}"" />
                                            </a:xfrm>
                                            <a:prstGeom prst=""rect"">
                                                <a:avLst />
                                            </a:prstGeom>
                                        </pic:spPr>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:inline>
                    </w:drawing></w:r>
                    ");

            return new Picture(document, xml, new Image(document, document.PackagePart.GetRelationship(id)));
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
                foreach (var el in textElement.Descendants())
                {
                    switch (el.Name.LocalName)
                    {
                        case "tab":
                            if (el.Parent?.Name.LocalName != "tabs")
                                goto case "br";
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
        /// Sets a specific text property with a value and child content
        /// </summary>
        /// <param name="propertyName">Text property name</param>
        /// <param name="value">Value (can be empty string)</param>
        /// <param name="content">Child content</param>
        internal void ApplyTextFormattingProperty(XName propertyName, string value, object content)
        {
            if (Runs.Count == 0)
            {
                var pPr = ParaElement();
                var rPr = pPr.GetRunProps();
                rPr.SetElementValue(propertyName, value);
                
                if (content is XAttribute attr)
                {
                    var last = rPr.Elements(propertyName).Last();
                    if (last.Attribute(attr.Name) == null)
                        last.Add(content);
                    else
                        last.SetAttributeValue(attr.Name, attr.Value);
                }
            }
            else
            {
                foreach (var run in Runs)
                {
                    var rPr = run.GetRunProps();
                    rPr.SetElementValue(propertyName, value);
                    var last = rPr.Elements(propertyName).Last();

                    switch (content)
                    {
                        case IEnumerable<object> properties:
                        {
                            foreach (var property in properties.Cast<XAttribute>())
                            {
                                var lastAttribute = last.Attribute(property.Name);
                                if (lastAttribute == null)
                                    last.Add(property);
                                else
                                    lastAttribute.Value = property.Value;
                            }

                            break;
                        }
                        case XAttribute attribute when last.Attribute(attribute.Name) == null:
                            last.Add(content);
                            break;
                        case XAttribute attribute:
                            last.SetAttributeValue(attribute.Name, attribute.Value);
                            break;
                        default:
                            if (content != null)
                                throw new NotSupportedException($"Unsupported content type {content.GetType().Name}: '{content}' for text formatting.");
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Walk all the text runs in the paragraph and find the one containing a specific index.
        /// </summary>
        /// <param name="index">Index to look for</param>
        /// <param name="editType">Type of edit being performed (insert or delete)</param>
        /// <returns>Run containing index</returns>
        internal Run GetFirstRunEffectedByEdit(int index, EditType editType = EditType.Ins)
        {
            int len = HelperFunctions.GetText(Xml).Length;
            if (index < 0 || editType == EditType.Ins && index > len || editType == EditType.Del && index >= len)
                throw new ArgumentOutOfRangeException(nameof(index));

            int count = 0;
            Run theOne = null;
            GetFirstRunEffectedByEditRecursive(Xml, index, ref count, ref theOne, editType);

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
        private void GetFirstRunEffectedByEditRecursive(XElement el, int index, ref int count, ref Run theOne, EditType editType)
        {
            count += HelperFunctions.GetSize(el);

            // If the EditType is deletion then we must return the next blah
            if (count > 0 && ((editType == EditType.Del && count > index) || (editType == EditType.Ins && count >= index)))
            {
                // Correct the index
                count = el.ElementsBeforeSelf().Aggregate(count, (current, e) => current - HelperFunctions.GetSize(e));
                count -= HelperFunctions.GetSize(el);

                // We have found the element, now find the run it belongs to.
                while (el != null && el.Name.LocalName != "r" && el.Name.LocalName != "pPr")
                {
                    el = el.Parent;
                }

                if (el == null)
                    throw new Exception($"Failed to locate index #{index} in paragraph.");

                theOne = new Run(Document, el, count);
            }
            else if (el.HasElements)
            {
                foreach (var e in el.Elements())
                {
                    if (theOne == null)
                    {
                        GetFirstRunEffectedByEditRecursive(e, index, ref count, ref theOne, editType);
                    }
                }
            }
        }

        /// <summary>
        /// If the pPr element doesn't exist it is created, either way it is returned by this function.
        /// </summary>
        /// <returns>The pPr element for this Paragraph.</returns>
        internal XElement ParaElement()
        {
            // Get the element.
            var pPr = Xml.Element(DocxNamespace.Main + "pPr");

            // If it doesn't exist, create it.
            if (pPr == null)
            {
                Xml.AddFirst(new XElement(DocxNamespace.Main + "pPr"));
                pPr = Xml.Element(DocxNamespace.Main + "pPr");
            }

            // Return the pPr element for this Paragraph.
            return pPr;
        }

        /// <summary>
        /// If the pPr/ind element doesn't exist it is created, either way it is returned by this function.
        /// </summary>
        /// <returns>The ind element for this Paragraphs pPr.</returns>
        internal XElement ParaIndentationElement()
        {
            // Get the element.
            var pPr = ParaElement();
            var ind = pPr.Element(DocxNamespace.Main + "ind");
            if (ind == null)
            {
                pPr.Add(new XElement(DocxNamespace.Main + "ind"));
                ind = pPr.Element(DocxNamespace.Main + "ind");
            }
            return ind;
        }

        /// <summary>
        /// Get or create a relationship link to a picture
        /// </summary>
        /// <param name="picture">Picture</param>
        /// <returns>Relationship id</returns>
        internal string GetOrCreateRelationship(Picture picture)
        {
            string uri = picture.Image.packageRelationship.TargetUri.OriginalString;

            // Search for a relationship with a TargetUri that points at this Image.
            string id = PackagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                .Where(r => r.TargetUri.OriginalString == uri)
                .Select(r => r.Id)
                .SingleOrDefault();

            if (id == null)
            {
                // Check to see if a relationship for this Picture exists and create it if not.
                var pr = PackagePart.CreateRelationship(picture.Image.packageRelationship.TargetUri, 
                    TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                id = pr.Id;
            }
            
            return id;
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

            foreach (var e in element.Elements())
            {
                if (RemoveHyperlinkRecursive(e, index, ref count))
                    return true;
            }
            
            return false;
        }

        /// <summary>
        /// Splits an XElement based on index
        /// </summary>
        /// <param name="element">Element to split</param>
        /// <param name="index">Index to split on</param>
        /// <param name="type">Type of edit being performed (insert/delete)</param>
        /// <returns>Split XElement array</returns>
        internal XElement[] SplitEdit(XElement element, int index, EditType type)
        {
            // Find the run containing the index
            var run = GetFirstRunEffectedByEdit(index, type);
            var splitRun = Run.SplitRun(run, index, type);

            var splitLeft = new XElement(element.Name, element.Attributes(), run.Xml.ElementsBeforeSelf(), splitRun[0]);
            if (GetElementTextLength(splitLeft) == 0)
            {
                splitLeft = null;
            }

            var splitRight = new XElement(element.Name, element.Attributes(), splitRun[1], run.Xml.ElementsAfterSelf());
            if (GetElementTextLength(splitRight) == 0)
            {
                splitRight = null;
            }

            return new[] { splitLeft, splitRight };
        }

    }
}