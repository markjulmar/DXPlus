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
    public abstract class Container : DocXElement
    {
        /// <summary>
        /// Returns a list of all Paragraphs inside this container.
        /// </summary>
        public virtual ReadOnlyCollection<Paragraph> Paragraphs
        {
            get
            {
                List<Paragraph> paragraphs = new List<Paragraph>();
                GetParagraphs(Xml, 0, paragraphs, false);

                foreach (Paragraph p in paragraphs)
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
            foreach (XElement paragraph in Xml.Descendants(DocxNamespace.Main + "p"))
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
            foreach (XElement paragraph in Xml.Descendants(DocxNamespace.Main + "p"))
            {
                if (paragraph.Equals(p.Xml))
                {
                    paragraph.Remove();
                    return true;
                }
            }

            return false;
        }

        public virtual List<Section> Sections
        {
            get
            {
                ReadOnlyCollection<Paragraph> allParas = Paragraphs;

                List<Paragraph> parasInASection = new List<Paragraph>();
                List<Section> sections = new List<Section>();

                foreach (Paragraph para in allParas)
                {
                    parasInASection.Add(para);

                    XElement sectionInPara = para.Xml.FirstLocalNameDescendant("sectPr");
                    if (sectionInPara != null)
                    {
                        sections.Add(new Section(Document, sectionInPara) { SectionParagraphs = parasInASection });
                        parasInASection = new List<Paragraph>();
                    }
                }

                XElement baseSectionXml = Xml.Element(DocxNamespace.Main + "body")?
                                             .Element(DocxNamespace.Main + "sectPr");
                if (baseSectionXml != null)
                {
                    sections.Add(new Section(Document, baseSectionXml) { SectionParagraphs = parasInASection });
                }

                return sections;
            }
        }

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
            XElement numNode = Document.numbering.LocalNameDescendants("num")?.FindByAttrVal(DocxNamespace.Main + "numId", numberDefRef);
            if (numNode != null)
            {
                string abstractNumNodeValue = numNode.FirstLocalNameDescendant("abstractNumId").GetVal();

                // Find the abstract numbering definition that defines the style of the numbering section.
                XElement abstractNumNode = Document.numbering.LocalNameDescendants("abstractNum")
                                                   .FindByAttrVal(DocxNamespace.Main + "abstractNumId", abstractNumNodeValue);

                // Get the numbering format.
                XElement numberingFormat = abstractNumNode.LocalNameDescendants("lvl")
                                            .FindByAttrVal(DocxNamespace.Main + "ilvl", numberingLevel)
                                            .FirstLocalNameDescendant("numFmt");

                p.ListItemType = numberingFormat.TryGetEnumValue<ListItemType>(out ListItemType result)
                    ? result
                    : ListItemType.Numbered;
            }
        }

        internal int GetParagraphs(XElement Xml, int index, List<Paragraph> paragraphs, bool deepSearch = false)
        {
            bool keepSearching = true;
            if (Xml.Name.LocalName == "p")
            {
                paragraphs.Add(SetParentContainerBasedOnType(new Paragraph(Document, Xml, index)));
                index += HelperFunctions.GetText(Xml).Length;
                if (!deepSearch)
                {
                    keepSearching = false;
                }
            }

            if (keepSearching && Xml.HasElements)
            {
                foreach (XElement e in Xml.Elements())
                {
                    index += GetParagraphs(e, index, paragraphs, deepSearch);
                }
            }

            return index;
        }

        public virtual IEnumerable<Table> Tables => Xml.Descendants(DocxNamespace.Main + "tbl")
                          .Select(t => new Table(Document, t) { packagePart = packagePart });

        public virtual List<List> Lists
        {
            get
            {
                List<List> lists = new List<List>();
                List list = new List(Document, Xml);

                foreach (Paragraph paragraph in Paragraphs)
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

        public virtual List<Hyperlink> Hyperlinks => Paragraphs.SelectMany(p => p.Hyperlinks).ToList();
        public virtual List<Picture> Pictures => Paragraphs.SelectMany(p => p.Pictures).ToList();

        /// <summary>
        /// Sets the Direction of content.
        /// </summary>
        /// <param name="direction">Direction either LeftToRight or RightToLeft</param>
        public virtual void SetDirection(Direction direction)
        {
            foreach (Paragraph p in Paragraphs)
            {
                p.Direction = direction;
            }
        }

        /// <summary>
        /// Find all occurances of a string in the paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <param name="ignoreCase"></param>
        /// <returns></returns>
        public virtual IEnumerable<int> FindAll(string text, bool ignoreCase = false)
        {
            foreach (Paragraph p in Paragraphs)
            {
                foreach (int index in p.FindAll(text, ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None))
                {
                    yield return index + p.startIndex;
                }
            }
        }

        /// <summary>
        /// Find all unique instances of the given Regex Pattern,
        /// returning the list of the unique strings found
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="options"></param>
        /// <returns>List of unique strings found.</returns>
        public virtual IEnumerable<(int index, string text)> FindPattern(string pattern, RegexOptions options)
        {
            foreach (Paragraph p in Paragraphs)
            {
                foreach ((int index, string text) in p.FindPattern(pattern, options))
                {
                    yield return (index: index + p.startIndex, text);
                }
            }
        }

        public virtual void ReplaceText(string searchValue, string newValue, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch, bool escapeRegEx = true, bool useRegExSubstitutions = false)
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
            foreach (Header header in new List<Header> { Document.Headers.First, Document.Headers.Even, Document.Headers.Odd }.Where(h => h != null))
            {
                foreach (Paragraph paragraph in header.Paragraphs)
                {
                    paragraph.ReplaceText(searchValue, newValue, trackChanges, options,
                        newFormatting, matchFormatting, formattingOptions,
                        escapeRegEx, useRegExSubstitutions);
                }
            }

            // ReplaceText int main body of document.
            foreach (Paragraph paragraph in Paragraphs)
            {
                paragraph.ReplaceText(searchValue, newValue, trackChanges, options,
                    newFormatting, matchFormatting, formattingOptions,
                    escapeRegEx, useRegExSubstitutions);
            }

            // ReplaceText in Footers of the document.
            foreach (Footer footer in new List<Footer> { Document.Footers.First, Document.Footers.Even, Document.Footers.Odd }.Where(h => h != null))
            {
                foreach (Paragraph paragraph in footer.Paragraphs)
                {
                    paragraph.ReplaceText(searchValue, newValue, trackChanges, options,
                        newFormatting, matchFormatting, formattingOptions,
                        escapeRegEx, useRegExSubstitutions);
                }
            }
        }

        /// <summary>
        /// Replace text based on a regex handler.
        /// </summary>
        /// <param name="searchValue">Value to find</param>
        /// <param name="regexMatchHandler">A Func that accepts the matching regex search group value and passes it to this to return the replacement string</param>
        /// <param name="trackChanges">Enable trackchanges</param>
        /// <param name="options">Regex options</param>
        /// <param name="newFormatting"></param>
        /// <param name="matchFormatting"></param>
        /// <param name="formattingOptions"></param>
        public virtual void ReplaceText(string searchValue, Func<string, string> regexMatchHandler, bool trackChanges = false, RegexOptions options = RegexOptions.None, Formatting newFormatting = null, Formatting matchFormatting = null, MatchFormattingOptions formattingOptions = MatchFormattingOptions.SubsetMatch)
        {
            if (string.IsNullOrEmpty(searchValue))
            {
                throw new ArgumentException("oldValue cannot be null or empty", nameof(searchValue));
            }

            if (regexMatchHandler == null)
            {
                throw new ArgumentException("regexMatchHandler cannot be null", nameof(regexMatchHandler));
            }

            // ReplaceText in Headers/Footers of the document.
            List<Container> containerList = new List<Container> {
                Document.Headers.First, Document.Headers.Even, Document.Headers.Odd,
                Document.Footers.First, Document.Footers.Even, Document.Footers.Odd
            };

            foreach (Container container in containerList.Where(c => c != null))
            {
                foreach (Paragraph paragraph in container.Paragraphs)
                {
                    paragraph.ReplaceText(searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, formattingOptions);
                }
            }

            // ReplaceText int main body of document.
            foreach (Paragraph paragraph in Paragraphs)
            {
                paragraph.ReplaceText(searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, formattingOptions);
            }
        }

        /// <summary>
        /// Removes all items with required formatting
        /// </summary>
        /// <returns>Numer of texts removed</returns>
        public int RemoveTextInGivenFormat(Formatting matchFormatting, MatchFormattingOptions matchOptions = MatchFormattingOptions.SubsetMatch)
        {
            int deletedCount = 0;
            foreach (XElement x in Xml.Elements())
            {
                deletedCount += RemoveTextWithFormatRecursive(x, matchFormatting, matchOptions);
            }

            return deletedCount;
        }

        internal int RemoveTextWithFormatRecursive(XElement element, Formatting matchFormatting, MatchFormattingOptions fo)
        {
            int deletedCount = 0;
            foreach (XElement x in element.Elements())
            {
                if ("rPr".Equals(x.Name.LocalName)
                    && HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, x, fo))
                {
                    x.Parent.Remove();
                    ++deletedCount;
                }

                deletedCount += RemoveTextWithFormatRecursive(x, matchFormatting, fo);
            }

            return deletedCount;
        }

        public virtual void InsertAtBookmark(string toInsert, string bookmarkName)
        {
            if (string.IsNullOrWhiteSpace(bookmarkName))
            {
                throw new ArgumentException("bookmark cannot be null or empty", nameof(bookmarkName));
            }

            Headers headerCollection = Document.Headers;
            List<Header> headers = new List<Header> { headerCollection.First, headerCollection.Even, headerCollection.Odd };
            foreach (Header header in headers.Where(x => x != null))
            {
                foreach (Paragraph paragraph in header.Paragraphs)
                {
                    paragraph.InsertAtBookmark(toInsert, bookmarkName);
                }
            }

            foreach (Paragraph paragraph in Paragraphs)
            {
                paragraph.InsertAtBookmark(toInsert, bookmarkName);
            }

            Footers footerCollection = Document.Footers;
            List<Footer> footers = new List<Footer> { footerCollection.First, footerCollection.Even, footerCollection.Odd };
            foreach (Footer footer in footers.Where(x => x != null))
            {
                foreach (Paragraph paragraph in footer.Paragraphs)
                {
                    paragraph.InsertAtBookmark(toInsert, bookmarkName);
                }
            }
        }

        public string[] ValidateBookmarks(params string[] bookmarkNames)
        {
            List<Header> headers = new[] { Document.Headers.First, Document.Headers.Even, Document.Headers.Odd }.Where(h => h != null).ToList();
            List<Footer> footers = new[] { Document.Footers.First, Document.Footers.Even, Document.Footers.Odd }.Where(f => f != null).ToList();

            List<string> nonMatching = new List<string>();
            foreach (string bookmarkName in bookmarkNames)
            {
                if (headers.SelectMany(h => h.Paragraphs).Any(p => p.ValidateBookmark(bookmarkName))
                    || footers.SelectMany(h => h.Paragraphs).Any(p => p.ValidateBookmark(bookmarkName))
                    || Paragraphs.Any(p => p.ValidateBookmark(bookmarkName)))
                {
                    return Array.Empty<string>();
                }

                nonMatching.Add(bookmarkName);
            }

            return nonMatching.ToArray();
        }

        public virtual Paragraph InsertParagraph(int index, Paragraph p)
        {
            XElement newXElement = new XElement(p.Xml);
            p.Xml = newXElement;

            Paragraph paragraph = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);
            if (paragraph == null)
            {
                Xml.Add(p.Xml);
            }
            else
            {
                XElement[] split = HelperFunctions.SplitParagraph(paragraph, index - paragraph.startIndex);
                paragraph.Xml.ReplaceWith(split[0], newXElement, split[1]);
            }

            return SetParentContainerBasedOnType(p);
        }

        public virtual Paragraph InsertParagraph(Paragraph p)
        {
            XDocument styleDoc;
            PackagePart stylePackagePart;

            if (p.styles.Count > 0)
            {
                Uri stylePackage = new Uri("/word/styles.xml", UriKind.Relative);
                if (!Document.package.PartExists(stylePackage))
                {
                    stylePackagePart = Document.package.CreatePart(stylePackage, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum);
                    using TextWriter tw = new StreamWriter(stylePackagePart.GetStream());
                    styleDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                                             new XElement(DocxNamespace.Main + "styles"));
                    styleDoc.Save(tw);
                }
                else
                {
                    stylePackagePart = Document.package.GetPart(stylePackage);
                    using TextReader tr = new StreamReader(stylePackagePart.GetStream());
                    styleDoc = XDocument.Load(tr);
                }

                // Get all the styleId values from the current style
                XElement styles = styleDoc.Element(DocxNamespace.Main + "styles");
                List<string> ids = styles.Descendants(DocxNamespace.Main + "style")
                                .Select(e => e.Attribute(DocxNamespace.Main + "styleId")?.Value)
                                .Where(v => v != null)
                                .ToList();

                // Go through the new paragraph and make sure all the styles are present
                foreach (XElement style in p.styles.Where(s => !ids.Contains(s.AttributeValue(DocxNamespace.Main + "styleId"))))
                {
                    styles.Add(style);
                }

                using (TextWriter tw = new StreamWriter(stylePackagePart.GetStream()))
                {
                    styleDoc.Save(tw);
                }
            }

            XElement newXElement = new XElement(p.Xml);

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

            Paragraph newParagraph = new Paragraph(Document, newXElement, index);
            Document.paragraphLookup.Add(index, newParagraph);
            SetParentContainerBasedOnType(newParagraph);
            return newParagraph;
        }

        public virtual Paragraph InsertParagraph(int index, string text, bool trackChanges, Formatting formatting)
        {
            Paragraph newParagraph = new Paragraph(Document, new XElement(DocxNamespace.Main + "p"), index);
            newParagraph.InsertText(0, text, trackChanges, formatting);

            Paragraph firstPar = HelperFunctions.GetFirstParagraphAffectedByInsert(Document, index);

            if (firstPar != null)
            {
                int splitindex = index - firstPar.startIndex;
                if (splitindex <= 0)
                {
                    firstPar.Xml.ReplaceWith(newParagraph.Xml, firstPar.Xml);
                }
                else
                {
                    XElement[] splitParagraph = HelperFunctions.SplitParagraph(firstPar, splitindex);

                    firstPar.Xml.ReplaceWith
                    (
                        splitParagraph[0],
                        newParagraph.Xml,
                        splitParagraph[1]
                    );
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
            XElement newParagraphSection = new XElement(DocxNamespace.Main + "p",
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
            XElement newParagraphSection = new XElement(DocxNamespace.Main + "p",
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