using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Base container object - this is used to represent all DOCX elements that contain child elements
    /// </summary>
    public abstract class Container : DocXBase, IContainer
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="document"></param>
        /// <param name="xml"></param>
        internal Container(IDocument document, XElement xml)
            : base(document, xml)
        {
        }

        /// <summary>
        /// Returns a list of all Paragraphs inside this container.
        /// </summary>
        public virtual ReadOnlyCollection<Paragraph> Paragraphs
        {
            get
            {
                var paragraphs = new List<Paragraph>();
                GetParagraphs(Xml, 0, paragraphs, false);

                foreach (var p in paragraphs)
                {
                    // If the next sibling node is a table, then link it up to this
                    // paragraph node.
                    var nextNode = p.Xml.ElementsAfterSelf().FirstOrDefault();
                    if (nextNode?.Name.Equals(Namespace.Main + "tbl") == true)
                    {
                        p.FollowingTable = new Table(Document, nextNode);
                    }

                    // Set the parent container type
                    p.ParentContainerType = p.Xml.Ancestors().First().Name.LocalName switch
                    {
                        "body" => ContainerType.Body,
                        "p" => ContainerType.Paragraph,
                        "tbl" => ContainerType.Table,
                        "sectPr" => ContainerType.Section,
                        "tc" => ContainerType.Cell,
                        _ => ContainerType.None,
                    };
                }

                return paragraphs.AsReadOnly();
            }
        }

        /// <summary>
        /// Removes paragraph at specified position
        /// </summary>
        /// <param name="index">Index of paragraph to remove</param>
        /// <returns>True if removed</returns>
        public bool RemoveParagraph(int index)
        {
            int i = 0;
            foreach (var paragraph in Xml.Descendants(Name.Paragraph))
            {
                if (i == index)
                {
                    paragraph.Remove();
                    return true;
                }
                ++i;
            }

            return false;
        }

        /// <summary>
        /// Removes paragraph
        /// </summary>
        /// <param name="paragraph">Paragraph to remove</param>
        /// <returns>True if removed</returns>
        public bool RemoveParagraph(Paragraph paragraph)
        {
            var paraXml = Xml.Descendants(Name.Paragraph).FirstOrDefault(p => p.Equals(paragraph.Xml));
            if (paraXml != null)
            {
                paraXml.Remove();
                return true;
            }
            return false;
        }

        /// <summary>
        /// Returns all the sections associated with this container.
        /// </summary>
        public List<Section> Sections
        {
            get
            {
                var parasInASection = new List<Paragraph>();
                var sections = new List<Section>();

                foreach (var para in Paragraphs)
                {
                    parasInASection.Add(para);

                    var sectionInPara = para.Xml.FirstLocalNameDescendant("sectPr");
                    if (sectionInPara != null)
                    {
                        sections.Add(new Section(Document, sectionInPara) { SectionParagraphs = parasInASection, PackagePart = PackagePart });
                        parasInASection = new List<Paragraph>();
                    }
                }

                var baseSectionXml = Xml.Element(Namespace.Main + "body")?
                                             .Element(Namespace.Main + "sectPr");
                if (baseSectionXml != null)
                {
                    sections.Add(new Section(Document, baseSectionXml) { SectionParagraphs = parasInASection, PackagePart = PackagePart });
                }

                return sections;
            }
        }


        /// <summary>
        /// Get all paragraphs in the document recursively.
        /// </summary>
        /// <param name="xml">XML to search</param>
        /// <param name="index">Max index</param>
        /// <param name="paragraphs">Found paragraphs</param>
        /// <param name="deepSearch">True to search inside paragraphs</param>
        /// <returns></returns>
        internal int GetParagraphs(XElement xml, int index, List<Paragraph> paragraphs, bool deepSearch = false)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml), "Paragraph collection not created.");

            bool keepSearching = true;
            if (xml.Name.LocalName == "p")
            {
                paragraphs.Add(SetParentContainerBasedOnType(new Paragraph(Document, xml, index)));
                index += HelperFunctions.GetText(xml).Length;
                if (!deepSearch)
                {
                    keepSearching = false;
                }
            }

            if (keepSearching && xml.HasElements)
            {
                index = xml.Elements().Aggregate(index, (current, e) => current + GetParagraphs(e, current, paragraphs, deepSearch));
            }

            return index;
        }

        /// <summary>
        /// Retrieve a list of all Table objects in the document
        /// </summary>
        public IEnumerable<Table> Tables => Xml.Descendants(Namespace.Main + "tbl")
                          .Select(t => new Table(Document, t));

        /// <summary>
        /// Retrieve a list of all Lists in the container (numbered or bulleted)
        /// </summary>
        public IEnumerable<List> Lists
        {
            get
            {
                var list = new List();
                foreach (var paragraph in Paragraphs)
                {
                    if (paragraph.IsListItem())
                    {
                        if (list.CanAddListItem(paragraph))
                        {
                            if (list.Items.Count == 0)
                            {
                                list.ListType = paragraph.GetNumberingFormat();
                                list.StartNumber = Document.NumberingStyles.GetStartingNumber(
                                    paragraph.GetListNumId(), paragraph.GetListLevel());
                            }

                            list.AddItem(paragraph);
                        }
                        // Found new list!
                        else
                        {
                            list.Document = Document;
                            list.PackagePart = PackagePart;
                            yield return list;

                            list = new List(paragraph.GetNumberingFormat(),
                                Document.NumberingStyles.GetStartingNumber(
                                    paragraph.GetListNumId(), paragraph.GetListLevel()));
                            list.AddItem(paragraph);
                        }
                    }
                }

                if (list.Items.Count > 0)
                {
                    list.Document = Document;
                    list.PackagePart = PackagePart;
                    yield return list;
                }
            }
        }

        /// <summary>
        /// Retrieve a list of all hyperlinks in the document
        /// </summary>
        public IEnumerable<Hyperlink> Hyperlinks => Paragraphs.SelectMany(p => p.Hyperlinks);

        /// <summary>
        /// Retrieve a list of all images (pictures) in the document
        /// </summary>
        public IEnumerable<Picture> Pictures => Paragraphs.SelectMany(p => p.Pictures);

        /// <summary>
        /// Sets the Direction of content.
        /// </summary>
        /// <param name="direction">Direction either LeftToRight or RightToLeft</param>
        public void SetDirection(Direction direction)
        {
            foreach (var p in Paragraphs)
            {
                p.Direction = direction;
            }
        }

        /// <summary>
        /// Replace matched text with a new value
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
                RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, 
                MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch, 
                bool escapeRegEx = true, bool useRegExSubstitutions = false)
        {
            if (string.IsNullOrEmpty(searchValue))
                throw new ArgumentNullException(nameof(searchValue));
            if (newValue == null)
                throw new ArgumentNullException(nameof(newValue));

            // ReplaceText in Headers of the document.
            if (Document != null)
            {
                foreach (var header in new List<Header> { Document.Headers.First, Document.Headers.Even, Document.Headers.Default })
                {
                    if (header.Exists)
                    {
                        foreach (var paragraph in header.Paragraphs)
                        {
                            paragraph.ReplaceText(searchValue, newValue, options,
                                newFormatting, matchFormatting, formattingOptions,
                                escapeRegEx, useRegExSubstitutions);
                        }
                    }
                }
            }

            // ReplaceText int main body of document.
            foreach (var paragraph in Paragraphs)
            {
                paragraph.ReplaceText(searchValue, newValue, options,
                    newFormatting, matchFormatting, formattingOptions,
                    escapeRegEx, useRegExSubstitutions);
            }

            // ReplaceText in Footers of the document.
            if (Document != null)
            {
                foreach (var footer in new List<Footer> { Document.Footers.First, Document.Footers.Even, Document.Footers.Default })
                {
                    if (footer.Exists)
                    {
                        foreach (var paragraph in footer.Paragraphs)
                        {
                            paragraph.ReplaceText(searchValue, newValue, options,
                                newFormatting, matchFormatting, formattingOptions,
                                escapeRegEx, useRegExSubstitutions);
                        }
                    }
                }
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
                throw new ArgumentException("bookmark cannot be null or empty", nameof(bookmarkName));

            // Try headers first
            if (Document != null)
            {
                var headerCollection = Document.Headers;
                var headers = new List<Header> {headerCollection.First, headerCollection.Even, headerCollection.Default};
                if (headers.Where(hdr => hdr.Exists).SelectMany(header => header.Paragraphs)
                    .Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert)))
                    return true;
            }

            // Body
            if (Paragraphs.Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert)))
                return true;

            // Footers
            if (Document != null)
            {
                var footerCollection = Document.Footers;
                var footers = new List<Footer> { footerCollection.First, footerCollection.Even, footerCollection.Default };
                return footers.Where(ftr => ftr.Exists).SelectMany(footer => footer.Paragraphs)
                    .Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert));
            }

            return false;
        }

        /// <summary>
        /// Insert a paragraph into this container at a specific index
        /// </summary>
        /// <param name="index">Character index to insert into</param>
        /// <param name="p">Paragraph to insert</param>
        /// <returns>Inserted paragraph</returns>
        public Paragraph InsertParagraph(int index, Paragraph p)
        {
            InsertMissingStyles(p);

            // Clone the element in case this is already in the document
            var newXElement = new XElement(p.Xml);
            p.Xml = newXElement;

            var paragraph = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);
            if (paragraph == null)
            {
                Xml.Add(p.Xml);
            }
            else
            {
                var split = HelperFunctions.SplitParagraph(paragraph, index - paragraph.StartIndex);
                paragraph.Xml.ReplaceWith(split[0], newXElement, split[1]);
            }

            return SetParentContainerBasedOnType(p);
        }

        internal Dictionary<int, Paragraph> GetParagraphIndexes() => Paragraphs.ToDictionary(paragraph => paragraph.EndIndex);

        /// <summary>
        /// Add a paragraph at the end of the container
        /// </summary>
        public Paragraph AddParagraph(Paragraph paragraph)
        {
            InsertMissingStyles(paragraph);

            // Clone the element in case this is already in the document
            var newXElement = new XElement(paragraph.Xml);
            Xml.Add(newXElement);

            var lookup = GetParagraphIndexes();
            int index = 0;
            if (lookup.Keys.Count > 0)
            {
                index = lookup.Last().Key;
                if (lookup.Last().Value.Text.Length == 0)
                {
                    index++;
                }
                else
                {
                    index += lookup.Last().Value.Text.Length;
                }
            }

            return SetParentContainerBasedOnType(new Paragraph(Document, newXElement, index));
        }

        /// <summary>
        /// Insert any missing styles associated with the passed paragraph
        /// </summary>
        /// <param name="p"></param>
        private void InsertMissingStyles(Paragraph p)
        {
            if (Document == null)
                return;

            // Make sure the document has all the styles associated to the
            // paragraph we are inserting.
            if (p.Styles.Count > 0)
            {
                var stylePackage = Relations.Styles.Uri;
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
                var styles = styleDoc.GetOrCreateElement(Namespace.Main + "styles");
                var ids = styles.Descendants(Namespace.Main + "style")
                                .Select(e => e.AttributeValue(Namespace.Main + "styleId", null))
                                .Where(v => v != null)
                                .ToList();

                // Go through the new paragraph and make sure all the styles are present
                foreach (var style in p.Styles.Where(s => !ids.Contains(s.AttributeValue(Namespace.Main + "styleId"))))
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
        public Paragraph InsertParagraph(int index, string text, Formatting formatting)
        {
            var newParagraph = new Paragraph(Document, new XElement(Name.Paragraph), index);
            newParagraph.InsertText(0, text, formatting);

            var firstPar = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);
            if (firstPar != null)
            {
                int splitIndex = index - firstPar.StartIndex;
                if (splitIndex <= 0)
                {
                    firstPar.Xml.ReplaceWith(newParagraph.Xml, firstPar.Xml);
                }
                else
                {
                    var splitParagraph = HelperFunctions.SplitParagraph(firstPar, splitIndex);
                    firstPar.Xml.ReplaceWith(splitParagraph[0], newParagraph.Xml, splitParagraph[1]);
                }
            }
            else
            {
                Xml.Add(newParagraph);
            }

            return SetParentContainerBasedOnType(newParagraph);
        }

        /// <summary>
        /// Determine the container parent type for the given paragraph
        /// </summary>
        /// <param name="paragraph">Paragraph</param>
        private Paragraph SetParentContainerBasedOnType(Paragraph paragraph)
        {
            paragraph.Container = this;
            paragraph.ParentContainerType = GetType().Name switch
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

            return paragraph;
        }

        /// <summary>
        /// Add a new section to the container
        /// </summary>
        public void AddSection() => Xml.Add(new XElement(Name.Paragraph,
                                                new XElement(Name.ParagraphProperties,
                                                    new XElement(Namespace.Main + "sectPr",
                                                        new XElement(Namespace.Main + "type",
                                                            new XAttribute(Name.MainVal, "continuous"))))));

        /// <summary>
        /// Add a new page break to the container
        /// </summary>
        public void AddPageBreak() => Xml.Add(new XElement(Name.Paragraph,
                                                new XElement(Name.ParagraphProperties,
                                                    new XElement(Namespace.Main + "sectPr"))));

        /// <summary>
        /// Add a paragraph with the given text to the end of the container
        /// </summary>
        /// <param name="text">Text to add</param>
        /// <param name="formatting">Formatting to use</param>
        /// <returns></returns>
        public Paragraph AddParagraph(string text, Formatting formatting)
        {
            if (text == null)
                throw new ArgumentNullException(nameof(text));

            var paragraph = ParagraphHelpers.Create(text, formatting);
            Xml.Add(paragraph);

            return SetParentContainerBasedOnType(new Paragraph(Document, paragraph, 0));
        }

        /// <summary>
        /// Add a new table to the end of the container
        /// </summary>
        /// <param name="table">Table to add</param>
        /// <returns>Table reference - may be copied if original table was already in document.</returns>
        public Table AddTable(Table table)
        {
            if (table.Document != null)
                table = new Table(Document, new XElement(table.Xml)) { Design = table.Design };

            table.Container = this;

            Xml.Add(table.Xml);

            return table;
        }

        /// <summary>
        /// Insert a Table into this document. The Table's source can be a completely different document.
        /// </summary>
        /// <param name="table">The Table to insert.</param>
        /// <param name="index">The index to insert this Table at.</param>
        /// <returns>The Table now associated with this document.</returns>
        public Table InsertTable(int index, Table table)
        {
            if (table.Document != null)
                table = new Table(Document, new XElement(table.Xml)) { Design = table.Design };

            table.Container = this;

            var firstParagraph = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);
            var split = HelperFunctions.SplitParagraph(firstParagraph, index - firstParagraph.StartIndex);
            firstParagraph.Xml.ReplaceWith(split[0], table.Xml, split[1]);

            return table;
        }

        /// <summary>
        /// Insert a List into this document.
        /// The List's source can be a completely different document.
        /// </summary>
        /// <param name="list">The List to insert</param>
        /// <returns>The List now associated with this document.</returns>
        public List AddList(List list)
        {
            if (list.Container != null)
                list = new List(list);

            foreach (var item in list.Items)
            {
                Xml.Add(item.Paragraph.Xml);
            }

            list.Container = this;

            return list;
        }

        /// <summary>
        /// Insert a List with a specific font size into this document.
        /// The List's source can be a completely different document.
        /// </summary>
        /// <param name="list">The List to insert</param>
        /// <param name="fontSize">Font size</param>
        /// <returns>The List now associated with this document.</returns>
        public List AddList(List list, double fontSize)
        {
            if (list.Container != null)
                list = new List(list);

            foreach (var item in list.Items)
            {
                item.Paragraph.FontSize(fontSize);
                Xml.Add(item.Paragraph.Xml);
            }

            list.Container = this;
            return list;
        }

        /// <summary>
        /// Insert a List with a specific font size/family into this document.
        /// The List's source can be a completely different document.
        /// </summary>
        /// <param name="list">The List to insert</param>
        /// <param name="fontFamily">Font family</param>
        /// <param name="fontSize">Font size</param>
        /// <returns>The List now associated with this document.</returns>
        public List AddList(List list, System.Drawing.FontFamily fontFamily, double fontSize)
        {
            if (list.Container != null)
                list = new List(list);

            foreach (var item in list.Items)
            {
                item.Paragraph.Font(fontFamily);
                item.Paragraph.FontSize(fontSize);
                Xml.Add(item.Paragraph.Xml);
            }

            list.Container = this;
            return list;
        }

        /// <summary>
        /// Insert a List at a specific position.
        /// </summary>
        /// <param name="index">Position to insert into</param>
        /// <param name="list">The List to insert</param>
        /// <returns>The List now associated with this document.</returns>
        public List InsertList(int index, List list)
        {
            if (list.Container != null)
                list = new List(list);

            list.Container = this;

            if (list.Items.Count > 0)
            {
                var firstParagraph = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);
                var split = HelperFunctions.SplitParagraph(firstParagraph, index - firstParagraph.StartIndex);

                var elements = new List<XElement> { split[0] };
                elements.AddRange(list.Items.Select(item => item.Xml));
                elements.Add(split[1]);
                firstParagraph.Xml.ReplaceWith(elements.ToArray<object>());
            }

            return list;
        }

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected override void OnDocumentOwnerChanged(IDocument previousValue, IDocument newValue)
        {
            base.OnDocumentOwnerChanged(previousValue, newValue);

            // Make sure we have all the styles in our document owner.
            if (newValue != null && Xml != null)
            {
                foreach (var paragraph in Paragraphs)
                {
                    InsertMissingStyles(paragraph);
                }
            }
        }
    }
}