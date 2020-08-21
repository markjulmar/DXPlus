using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
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
    public abstract class Container : DocXElement
    {
        internal Container(DocX document, XElement xml)
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
                    if (nextNode?.Name.Equals(DocxNamespace.Main + "tbl") == true)
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
        public bool RemoveParagraphAt(int index)
        {
            int i = 0;
            foreach (var paragraph in Xml.Descendants(DocxNamespace.Main + "p"))
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
        /// <param name="p">Paragraph to remove</param>
        /// <returns>True if removed</returns>
        public bool RemoveParagraph(Paragraph p)
        {
            foreach (var paragraph in Xml.Descendants(DocxNamespace.Main + "p"))
            {
                if (paragraph.Equals(p.Xml))
                {
                    paragraph.Remove();
                    return true;
                }
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
                        sections.Add(new Section(Document, sectionInPara) { SectionParagraphs = parasInASection });
                        parasInASection = new List<Paragraph>();
                    }
                }

                var baseSectionXml = Xml.Element(DocxNamespace.Main + "body")?
                                             .Element(DocxNamespace.Main + "sectPr");
                if (baseSectionXml != null)
                {
                    sections.Add(new Section(Document, baseSectionXml) { SectionParagraphs = parasInASection });
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
        public IEnumerable<Table> Tables => Xml.Descendants(DocxNamespace.Main + "tbl")
                          .Select(t => new Table(Document, t) { PackagePart = PackagePart });

        /// <summary>
        /// Retrieve a list of all Lists in the container (numbered or bulleted)
        /// </summary>
        public List<List> Lists
        {
            get
            {
                var lists = new List<List>();
                var list = new List {Document = Document, Xml = Xml};

                foreach (var paragraph in Paragraphs)
                {
                    if (paragraph.IsListItem)
                    {
                        if (list.CanAddListItem(paragraph))
                        {
                            list.AddItem(paragraph);
                        }
                        else
                        {
                            lists.Add(list);
                            list = new List {Document = Document, Xml = Xml};
                            list.AddItem(paragraph);
                        }
                    }
                }

                lists.Add(list);

                return lists;
            }
        }

        /// <summary>
        /// Retrieve a list of all hyperlinks in the document
        /// </summary>
        public List<Hyperlink> Hyperlinks => Paragraphs.SelectMany(p => p.Hyperlinks).ToList();

        /// <summary>
        /// Retrieve a list of all images (pictures) in the document
        /// </summary>
        public List<Picture> Pictures => Paragraphs.SelectMany(p => p.Pictures).ToList();

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
        /// Find all occurrences of a string in the paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <param name="ignoreCase"></param>
        /// <returns></returns>
        public IEnumerable<int> FindAll(string text, bool ignoreCase = false)
        {
            return from p in Paragraphs
                   from index in p.FindAll(text, ignoreCase) 
                   select index + p.StartIndex;
        }

        /// <summary>
        /// Find all unique instances of the given Regex Pattern,
        /// returning the list of the unique strings found
        /// </summary>
        /// <param name="regex">Pattern to search for</param>
        /// <returns>Index and matched strings</returns>
        public IEnumerable<(int index, string text)> FindPattern(Regex regex)
        {
            foreach (var p in Paragraphs)
            {
                foreach ((int index, string text) in p.FindPattern(regex))
                {
                    yield return (index: index + p.StartIndex, text);
                }
            }
        }

        /// <summary>
        /// Replace matched text with a new value
        /// </summary>
        /// <param name="searchValue">Text value to search for</param>
        /// <param name="newValue">Replacement value</param>
        /// <param name="trackChanges">True to track changes</param>
        /// <param name="options">Regex options</param>
        /// <param name="newFormatting">New formatting to apply</param>
        /// <param name="matchFormatting">Formatting to match</param>
        /// <param name="formattingOptions">Match formatting options</param>
        /// <param name="escapeRegEx">True to escape Regex expression</param>
        /// <param name="useRegExSubstitutions">True to use RegEx in substitution</param>
        public void ReplaceText(string searchValue, string newValue, bool trackChanges = false, 
                RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, 
                MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch, 
                bool escapeRegEx = true, bool useRegExSubstitutions = false)
        {
            if (string.IsNullOrEmpty(searchValue))
            {
                throw new ArgumentException("oldValue cannot be null or empty", nameof(searchValue));
            }

            if (newValue == null)
            {
                throw new ArgumentException("newValue cannot be null or empty", nameof(newValue));
            }

            // ReplaceText in Headers of the document.
            foreach (var header in new List<Header> { Document.Headers.First, Document.Headers.Even, Document.Headers.Odd }.Where(h => h != null))
            {
                foreach (var paragraph in header.Paragraphs)
                {
                    paragraph.ReplaceText(searchValue, newValue, trackChanges, options,
                        newFormatting, matchFormatting, formattingOptions,
                        escapeRegEx, useRegExSubstitutions);
                }
            }

            // ReplaceText int main body of document.
            foreach (var paragraph in Paragraphs)
            {
                paragraph.ReplaceText(searchValue, newValue, trackChanges, options,
                    newFormatting, matchFormatting, formattingOptions,
                    escapeRegEx, useRegExSubstitutions);
            }

            // ReplaceText in Footers of the document.
            foreach (var footer in new List<Footer> { Document.Footers.First, Document.Footers.Even, Document.Footers.Odd }.Where(h => h != null))
            {
                foreach (var paragraph in footer.Paragraphs)
                {
                    paragraph.ReplaceText(searchValue, newValue, trackChanges, options,
                        newFormatting, matchFormatting, formattingOptions,
                        escapeRegEx, useRegExSubstitutions);
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
            var headerCollection = Document.Headers;
            var headers = new List<Header> { headerCollection.First, headerCollection.Even, headerCollection.Odd };
            if (headers.Where(x => x != null).SelectMany(header => header.Paragraphs).Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert)))
                return true;

            // Body
            if (Paragraphs.Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert)))
                return true;

            // Footers
            var footerCollection = Document.Footers;
            var footers = new List<Footer> { footerCollection.First, footerCollection.Even, footerCollection.Odd };
            return footers.Where(x => x != null).SelectMany(footer => footer.Paragraphs).Any(paragraph => paragraph.InsertAtBookmark(bookmarkName, toInsert));
        }

        /// <summary>
        /// Insert a paragraph into this container at a specific index
        /// </summary>
        /// <param name="index">Index to insert into</param>
        /// <param name="p">New paragraph</param>
        /// <returns></returns>
        public Paragraph InsertParagraph(int index, Paragraph p)
        {
            InsertMissingStyles(p);

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

        /// <summary>
        /// Insert a new paragraph at the end of the container
        /// </summary>
        /// <param name="p">New paragraph</param>
        /// <returns></returns>
        public Paragraph InsertParagraph(Paragraph p)
        {
            InsertMissingStyles(p);

            var newXElement = new XElement(p.Xml);
            Xml.Add(newXElement);

            int index = 0;
            if (Document.paragraphLookup.Keys.Count > 0)
            {
                index = Document.paragraphLookup.Last().Key;
                if (Document.paragraphLookup.Last().Value.Text.Length == 0)
                {
                    index++;
                }
                else
                {
                    index += Document.paragraphLookup.Last().Value.Text.Length;
                }
            }

            var newParagraph = new Paragraph(Document, newXElement, index);
            Document.paragraphLookup.Add(index, newParagraph);
            return SetParentContainerBasedOnType(newParagraph);
        }

        /// <summary>
        /// Insert any missing styles associated with the passed paragraph
        /// </summary>
        /// <param name="p"></param>
        private void InsertMissingStyles(Paragraph p)
        {
            // Make sure the document has all the styles associated to the
            // paragraph we are inserting.
            if (p.Styles.Count > 0)
            {
                Uri stylePackage = new Uri("/word/styles.xml", UriKind.Relative);
                XDocument styleDoc;
                PackagePart stylePackagePart;
                if (!Document.Package.PartExists(stylePackage))
                {
                    stylePackagePart = Document.Package.CreatePart(stylePackage, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum);
                    styleDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(DocxNamespace.Main + "styles"));
                    stylePackagePart.Save(styleDoc);
                }
                else
                {
                    stylePackagePart = Document.Package.GetPart(stylePackage);
                    styleDoc = stylePackagePart.Load();
                }

                // Get all the styleId values from the current style
                var styles = styleDoc.GetOrCreateElement(DocxNamespace.Main + "styles");
                var ids = styles.Descendants(DocxNamespace.Main + "style")
                                .Select(e => e.AttributeValue(DocxNamespace.Main + "styleId", null))
                                .Where(v => v != null)
                                .ToList();

                // Go through the new paragraph and make sure all the styles are present
                foreach (var style in p.Styles.Where(s => !ids.Contains(s.AttributeValue(DocxNamespace.Main + "styleId"))))
                {
                    styles.Add(style);
                }

                // Save back to the package
                stylePackagePart.Save(styleDoc);
            }

        }

        public Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
        {
            var newParagraph = new Paragraph(Document, new XElement(DocxNamespace.Main + "p"), index);
            newParagraph.InsertText(0, text, trackChanges, formatting);

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

        private Paragraph SetParentContainerBasedOnType(Paragraph newParagraph)
        {
            newParagraph.ParentContainerType = GetType().Name switch
            {
                nameof(Table) => ContainerType.Table,
                nameof(TableOfContents) => ContainerType.TOC,
                nameof(Section) => ContainerType.Section,
                nameof(Cell) => ContainerType.Cell,
                nameof(Header) => ContainerType.Header,
                nameof(Footer) => ContainerType.Footer,
                nameof(Paragraph) => ContainerType.Paragraph,
                _ => ContainerType.None
            };

            newParagraph.PackagePart = this switch
            {
                Cell cell => cell.PackagePart,
                DocX _ => Document.PackagePart,
                Footer f => f.PackagePart,
                Header h => h.PackagePart,
                Row r => r.PackagePart,
                _ => Document.PackagePart,
            };

            return newParagraph;
        }

        public void InsertSection(bool trackChanges = false)
        {
            var newParagraphSection = new XElement(DocxNamespace.Main + "p",
                new XElement(DocxNamespace.Main + "pPr",
                    new XElement(DocxNamespace.Main + "sectPr",
                        new XElement(DocxNamespace.Main + "type",
                            new XAttribute(DocxNamespace.Main + "val", "continuous")
                        )
                    )
                )
            );

            if (trackChanges)
            {
                newParagraphSection = HelperFunctions.CreateEdit(EditType.Ins, DateTime.Now, newParagraphSection);
            }

            Xml.Add(newParagraphSection);
        }

        public void InsertSectionPageBreak(bool trackChanges = false)
        {
            var newParagraphSection = new XElement(DocxNamespace.Main + "p",
                new XElement(DocxNamespace.Main + "pPr",
                    new XElement(DocxNamespace.Main + "sectPr")
                )
            );

            if (trackChanges)
            {
                newParagraphSection = HelperFunctions.CreateEdit(EditType.Ins, DateTime.Now, newParagraphSection);
            }

            Xml.Add(newParagraphSection);
        }

        public Paragraph InsertParagraph() => InsertParagraph(string.Empty);

        public Paragraph InsertParagraph(string text) => InsertParagraph(text, false);

        public Paragraph InsertParagraph(string text, bool trackChanges) => InsertParagraph(text, trackChanges, null);

        public Paragraph InsertParagraph(int index, string text, bool trackChanges) => InsertParagraph(index, text, trackChanges, null);
        
        public Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
        {
            if (text == null)
                throw new ArgumentNullException(nameof(text));

            var newParagraph = new XElement(DocxNamespace.Main + "p",
                                            HelperFunctions.FormatInput(text, formatting?.Xml));

            if (trackChanges)
            {
                newParagraph = HelperFunctions.CreateEdit(EditType.Ins, DateTime.Now, newParagraph);
            }

            Xml.Add(newParagraph);
            return SetParentContainerBasedOnType(new Paragraph(Document, newParagraph, 0));
        }

        public Paragraph InsertEquation(string equation)
        {
            return InsertParagraph()
                  .AppendEquation(equation);
        }

        public Paragraph InsertBookmark(string bookmarkName)
        {
            return InsertParagraph()
                  .AppendBookmark(bookmarkName);
        }

        public Table InsertTable(int rowCount, int columnCount)
        {
            return InsertTable(Document.CreateTable(rowCount, columnCount));
        }

        public Table InsertTable(Table t)
        {
            var newXElement = new XElement(t.Xml);
            Xml.Add(newXElement);

            return new Table(Document, newXElement)
            {
                PackagePart = PackagePart,
                Design = t.Design
            };
        }

        /// <summary>
        /// Insert a Table into this document. The Table's source can be a completely different document.
        /// </summary>
        /// <param name="t">The Table to insert.</param>
        /// <param name="index">The index to insert this Table at.</param>
        /// <returns>The Table now associated with this document.</returns>
        public Table InsertTable(int index, Table t)
        {
            var p = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);

            var split = HelperFunctions.SplitParagraph(p, index - p.StartIndex);
            var newXElement = new XElement(t.Xml);
            p.Xml.ReplaceWith(split[0], newXElement, split[1]);

            return new Table(Document, newXElement) { PackagePart = PackagePart, Design = t.Design };
        }

        /// <summary>
        /// Insert a List into this document.
        /// The List's source can be a completely different document.
        /// </summary>
        /// <param name="list">The List to insert</param>
        /// <returns>The List now associated with this document.</returns>
        public List InsertList(List list)
        {
            foreach (var item in list.Items)
            {
                Xml.Add(item.Xml);
            }

            return list;
        }

        /// <summary>
        /// Insert a List with a specific font size into this document.
        /// The List's source can be a completely different document.
        /// </summary>
        /// <param name="list">The List to insert</param>
        /// <param name="fontSize">Font size</param>
        /// <returns>The List now associated with this document.</returns>
        public List InsertList(List list, double fontSize)
        {
            foreach (var item in list.Items)
            {
                item.FontSize(fontSize);
                Xml.Add(item.Xml);
            }

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
        public List InsertList(List list, System.Drawing.FontFamily fontFamily, double fontSize)
        {
            foreach (var item in list.Items)
            {
                item.Font(fontFamily);
                item.FontSize(fontSize);
                Xml.Add(item.Xml);
            }

            return list;
        }

        /// <summary>
        /// Insert a List at a specific position. The List's source can be a completely different document.
        /// </summary>
        /// <param name="index">Position to insert into</param>
        /// <param name="list">The List to insert</param>
        /// <returns>The List now associated with this document.</returns>
        public List InsertList(int index, List list)
        {
            var p = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);

            var split = HelperFunctions.SplitParagraph(p, index - p.StartIndex);
            var elements = new List<XElement> { split[0] };
            elements.AddRange(list.Items.Select(i => new XElement(i.Xml)));
            elements.Add(split[1]);
            p.Xml.ReplaceWith(elements.ToArray<object>());

            return list;
        }
    }
}