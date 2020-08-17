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
                    XElement nextNode = p.Xml.ElementsAfterSelf().FirstOrDefault();
                    if (nextNode?.Name.Equals(DocxNamespace.Main + "tbl") == true)
                    {
                        p.FollowingTable = new Table(Document, nextNode);
                    }

                    // Set the parent container type
                    p.ParentContainer = p.Xml.Ancestors().First().Name.LocalName switch
                    {
                        "body" => ContainerType.Body,
                        "p" => ContainerType.Paragraph,
                        "tbl" => ContainerType.Table,
                        "sectPr" => ContainerType.Section,
                        "tc" => ContainerType.Cell,
                        _ => ContainerType.None,
                    };

                    if (p.IsListItem)
                    {
                        GetListItemType(p);
                    }
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
        public virtual List<Section> Sections
        {
            get
            {
                var parasInASection = new List<Paragraph>();
                var sections = new List<Section>();

                foreach (Paragraph para in Paragraphs)
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
        /// Set the ListItemType property on the passed paragraph based on the List type identified.
        /// Defaults to numbered if a list is found but the type is not specified
        /// </summary>
        /// <param name="p">Paragraph to check</param>
        private void GetListItemType(Paragraph p)
        {
            // <w:p>
            //   <w:pPr>
            //     <w:numPr>
            //       <w:ilvl w:val="0"/>
            //       <w:numId w:val="5"/>
            //     </w:numPr>
            //   </w:pPr>
            //</w:p>

            string numberingLevel = p.ParagraphNumberProperties.FirstLocalNameDescendant("ilvl").GetVal();
            string numberDefRef = p.ParagraphNumberProperties.FirstLocalNameDescendant("numId").GetVal();

            // Find the number definition instance.
            var numNode = Document.numbering.LocalNameDescendants("num")?.FindByAttrVal(DocxNamespace.Main + "numId", numberDefRef);
            if (numNode != null)
            {
                string abstractNumNodeValue = numNode.FirstLocalNameDescendant("abstractNumId").GetVal();

                // Find the abstract numbering definition that defines the style of the numbering section.
                var abstractNumNode = Document.numbering.LocalNameDescendants("abstractNum")
                                                   .FindByAttrVal(DocxNamespace.Main + "abstractNumId", abstractNumNodeValue);

                // Get the numbering format.
                var numberingFormat = abstractNumNode.LocalNameDescendants("lvl")
                                            .FindByAttrVal(DocxNamespace.Main + "ilvl", numberingLevel)
                                            .FirstLocalNameDescendant("numFmt");

                p.ListItemType = numberingFormat.TryGetEnumValue<ListItemType>(out ListItemType result)
                    ? result
                    : ListItemType.Numbered;
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
        public virtual IEnumerable<Table> Tables => Xml.Descendants(DocxNamespace.Main + "tbl")
                          .Select(t => new Table(Document, t) { packagePart = packagePart });

        /// <summary>
        /// Retrieve a list of all Lists in the document (numbered or bulleted)
        /// </summary>
        public virtual List<List> Lists
        {
            get
            {
                var lists = new List<List>();
                var list = new List(Document, Xml);

                foreach (var paragraph in Paragraphs)
                {
                    paragraph.packagePart = packagePart;
                    if (paragraph.IsListItem)
                    {
                        if (list.CanAddListItem(paragraph))
                        {
                            list.AddItem(paragraph);
                        }
                        else
                        {
                            lists.Add(list);
                            list = new List(Document, Xml);
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
        public virtual List<Hyperlink> Hyperlinks => Paragraphs.SelectMany(p => p.Hyperlinks).ToList();

        /// <summary>
        /// Retrieve a list of all images (pictures) in the document
        /// </summary>
        public virtual List<Picture> Pictures => Paragraphs.SelectMany(p => p.Pictures).ToList();

        /// <summary>
        /// Sets the Direction of content.
        /// </summary>
        /// <param name="direction">Direction either LeftToRight or RightToLeft</param>
        public virtual void SetDirection(Direction direction)
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
        public virtual IEnumerable<int> FindAll(string text, bool ignoreCase = false)
        {
            return from p in Paragraphs
                   from index in p.FindAll(text, ignoreCase) 
                   select index + p.startIndex;
        }

        /// <summary>
        /// Find all unique instances of the given Regex Pattern,
        /// returning the list of the unique strings found
        /// </summary>
        /// <param name="regex">Pattern to search for</param>
        /// <returns>Index and matched strings</returns>
        public virtual IEnumerable<(int index, string text)> FindPattern(Regex regex)
        {
            foreach (var p in Paragraphs)
            {
                foreach ((int index, string text) in p.FindPattern(regex))
                {
                    yield return (index: index + p.startIndex, text);
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
        public virtual void ReplaceText(string searchValue, string newValue, bool trackChanges = false, 
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
        public virtual bool InsertAtBookmark(string bookmarkName, string toInsert)
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
        public virtual Paragraph InsertParagraph(int index, Paragraph p)
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
                var split = HelperFunctions.SplitParagraph(paragraph, index - paragraph.startIndex);
                paragraph.Xml.ReplaceWith(split[0], newXElement, split[1]);
            }

            return SetParentContainerBasedOnType(p);
        }

        /// <summary>
        /// Insert a new paragraph at the end of the container
        /// </summary>
        /// <param name="p">New paragraph</param>
        /// <returns></returns>
        public virtual Paragraph InsertParagraph(Paragraph p)
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
            if (p.styles.Count > 0)
            {
                Uri stylePackage = new Uri("/word/styles.xml", UriKind.Relative);
                XDocument styleDoc;
                PackagePart stylePackagePart;
                if (!Document.Package.PartExists(stylePackage))
                {
                    stylePackagePart = Document.Package.CreatePart(stylePackage, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum);
                    using TextWriter tw = new StreamWriter(stylePackagePart.GetStream());
                    styleDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                                             new XElement(DocxNamespace.Main + "styles"));
                    styleDoc.Save(tw);
                }
                else
                {
                    stylePackagePart = Document.Package.GetPart(stylePackage);
                    using TextReader tr = new StreamReader(stylePackagePart.GetStream());
                    styleDoc = XDocument.Load(tr);
                }

                // Get all the styleId values from the current style
                XElement styles = styleDoc.GetOrCreateElement(DocxNamespace.Main + "styles");
                var ids = styles.Descendants(DocxNamespace.Main + "style")
                                .Select(e => e.AttributeValue(DocxNamespace.Main + "styleId", null))
                                .Where(v => v != null)
                                .ToList();

                // Go through the new paragraph and make sure all the styles are present
                foreach (var style in p.styles.Where(s => !ids.Contains(s.AttributeValue(DocxNamespace.Main + "styleId"))))
                {
                    styles.Add(style);
                }

                using (TextWriter tw = new StreamWriter(stylePackagePart.GetStream()))
                {
                    styleDoc.Save(tw);
                }
            }

        }

        public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
        {
            Paragraph newParagraph = new Paragraph(Document, new XElement(DocxNamespace.Main + "p"), index);
            newParagraph.InsertText(0, text, trackChanges, formatting);

            Paragraph firstPar = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);

            if (firstPar != null)
            {
                int splitIndex = index - firstPar.startIndex;
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
            newParagraph.ParentContainer = GetType().Name switch
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

            newParagraph.packagePart = this switch
            {
                Cell cell => cell.packagePart,
                DocX _ => Document.packagePart,
                Footer f => f.packagePart,
                Header h => h.packagePart,
                _ => Document.packagePart,
            };

            return newParagraph;
        }

        public virtual void InsertSection(bool trackChanges = false)
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

        public virtual void InsertSectionPageBreak(bool trackChanges = false)
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

        public virtual Paragraph InsertParagraph()
        {
            return InsertParagraph(string.Empty, false, new Formatting());
        }

        public virtual Paragraph InsertParagraph(string text)
        {
            return InsertParagraph(text, false, new Formatting());
        }

        public virtual Paragraph InsertParagraph(string text, bool trackChanges)
        {
            return InsertParagraph(text, trackChanges, new Formatting());
        }

        public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges)
        {
            return InsertParagraph(index, text, trackChanges, null);
        }

        public virtual Paragraph InsertParagraph(string text, bool trackChanges, Formatting formatting)
        {
            XElement newParagraph = new XElement(DocxNamespace.Main + "p",
                                        new XElement(DocxNamespace.Main + "pPr"),
                                            HelperFunctions.FormatInput(text, formatting.Xml));

            if (trackChanges)
            {
                newParagraph = HelperFunctions.CreateEdit(EditType.Ins, DateTime.Now, newParagraph);
            }

            Xml.Add(newParagraph);
            return SetParentContainerBasedOnType(new Paragraph(Document, newParagraph, 0));
        }

        public virtual Paragraph InsertEquation(string equation)
        {
            Paragraph p = InsertParagraph();
            p.AppendEquation(equation);
            return p;
        }

        public virtual Paragraph InsertBookmark(string bookmarkName)
        {
            Paragraph p = InsertParagraph();
            p.AppendBookmark(bookmarkName);
            return p;
        }

        public Table InsertTable(int rowCount, int columnCount)
        {
            return InsertTable(Document.CreateTable(rowCount, columnCount));
        }

        public Table InsertTable(Table t)
        {
            XElement newXElement = new XElement(t.Xml);
            Xml.Add(newXElement);

            return new Table(Document, newXElement)
            {
                packagePart = packagePart,
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
            Paragraph p = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);

            XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);
            XElement newXElement = new XElement(t.Xml);
            p.Xml.ReplaceWith(split[0], newXElement, split[1]);

            return new Table(Document, newXElement) { packagePart = packagePart, Design = t.Design };
        }

        internal Container(DocX document, XElement xml)
            : base(document, xml)
        {
        }

        public List InsertList(List list)
        {
            foreach (Paragraph item in list.Items)
            {
                Xml.Add(item.Xml);
            }

            return list;
        }

        public List InsertList(List list, double fontSize)
        {
            foreach (Paragraph item in list.Items)
            {
                item.FontSize(fontSize);
                Xml.Add(item.Xml);
            }

            return list;
        }

        public List InsertList(List list, System.Drawing.FontFamily fontFamily, double fontSize)
        {
            foreach (Paragraph item in list.Items)
            {
                item.Font(fontFamily);
                item.FontSize(fontSize);
                Xml.Add(item.Xml);
            }

            return list;
        }

        public List InsertList(int index, List list)
        {
            Paragraph p = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);

            XElement[] split = HelperFunctions.SplitParagraph(p, index - p.startIndex);
            List<XElement> elements = new List<XElement> { split[0] };
            elements.AddRange(list.Items.Select(i => new XElement(i.Xml)));
            elements.Add(split[1]);
            p.Xml.ReplaceWith(elements.ToArray());

            return list;
        }
    }
}