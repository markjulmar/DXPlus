using DXPlus.Helpers;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Base container object - this is used to represent all DOCX elements that contain child elements
    /// </summary>
    public abstract class BlockContainer : DocXElement, IContainer
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="document"></param>
        /// <param name="xml"></param>
        internal BlockContainer(IDocument document, XElement xml)
            : base(document, xml)
        {
        }

        /// <summary>
        /// Returns a list of all Paragraphs inside this container.
        /// </summary>
        public IEnumerable<Paragraph> Paragraphs
        {
            get
            {
                if (Xml != null)
                {
                    int current = 0;
                    foreach (var e in Xml.Elements(Name.Paragraph))
                    {
                        yield return HelperFunctions.WrapParagraphElement(e, Document, PackagePart, ref current);
                    }
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
            if (paragraph.BlockContainer == this)
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
                foreach (var para in Paragraphs)
                {
                    if (para.Xml.Element(Name.ParagraphProperties, Name.SectionProperties) != null)
                        yield return new Section(Document, para.Xml) {PackagePart = PackagePart};
                }

                // Return the final section if this is the mainDoc.
                if (Xml.Element(Name.SectionProperties) != null)
                {
                    yield return new Section(Document, Xml) {PackagePart = PackagePart}; 
                }
            }
        }

        /// <summary>
        /// Retrieve a list of all Table objects in the document
        /// </summary>
        public IEnumerable<Table> Tables => Xml.Descendants(Name.Table)
                          .Select(t => new Table(Document, t));

        /// <summary>
        /// Retrieve a list of all hyperlinks in the document
        /// </summary>
        public IEnumerable<Hyperlink> Hyperlinks => Paragraphs.SelectMany(p => p.Hyperlinks);

        /// <summary>
        /// Retrieve a list of all images (pictures) in the document
        /// </summary>
        public IEnumerable<Picture> Pictures => Paragraphs.SelectMany(p => p.Pictures);

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
            if (paragraph.InDom)
                throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

            Paragraph insertPos = Document.FindParagraphByIndex(index);
            if (insertPos == null)
            {
                AddElementToDocument(paragraph.Xml);
            }
            else
            {
                XElement[] split = SplitParagraph(insertPos, index - insertPos.StartIndex);
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

            if (paragraph.InDom)
                throw new ArgumentException("Cannot add paragraph multiple times.", nameof(paragraph));

            AddElementToDocument(paragraph.Xml);
            return OnAddParagraph(paragraph);
        }

        /// <summary>
        /// This method is called when a new paragraph is added to this container.
        /// </summary>
        /// <param name="paragraph">New paragraph</param>
        /// <returns>Added paragraph</returns>
        private Paragraph OnAddParagraph(Paragraph paragraph)
        {
            InsertMissingStyles(paragraph);

            paragraph.BlockContainer = this;
            paragraph.SetStartIndex(Paragraphs.Single(p => p.Id == paragraph.Id).StartIndex);

            return paragraph;
        }

        /// <summary>
        /// Adds a new paragraph into the document structure
        /// </summary>
        /// <param name="xml"></param>
        private void AddElementToDocument(XElement xml)
        {
            // On paragraphs, add an ID if it's missing.
            if (xml.Name.LocalName == "p" && xml.Attribute(Name.ParagraphId) == null)
            {
                xml.SetAttributeValue(Name.ParagraphId, HelperFunctions.GenerateHexId());
            }

            XElement sectPr = Xml.Elements(Name.SectionProperties).SingleOrDefault();
            if (sectPr != null)
            {
                sectPr.AddBeforeSelf(xml);
            }
            else
            {
                Xml.Add(xml);
            }
        }

        /// <summary>
        /// Insert any missing styles associated with the passed paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        private void InsertMissingStyles(Paragraph paragraph)
        {
            if (Document == null)
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
                                .ToList();

                // Go through the new paragraph and make sure all the styles are present
                foreach (XElement style in paragraph.Styles.Where(s => !ids.Contains(s.AttributeValue(Namespace.Main + "styleId"))))
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
            Paragraph newParagraph = new Paragraph(Document, Paragraph.Create(text, formatting), index);
            Paragraph firstPar = Document.FindParagraphByIndex(index);
            if (firstPar != null)
            {
                int splitIndex = index - firstPar.StartIndex;
                if (splitIndex <= 0)
                {
                    firstPar.Xml.ReplaceWith(newParagraph.Xml, firstPar.Xml);
                }
                else
                {
                    XElement[] splitParagraph = SplitParagraph(firstPar, splitIndex);
                    firstPar.Xml.ReplaceWith(splitParagraph[0], newParagraph.Xml, splitParagraph[1]);
                }
            }
            else
            {
                AddElementToDocument(newParagraph.Xml);
            }

            return OnAddParagraph(newParagraph);
        }

        /// <summary>
        /// Add a new section to the container
        /// </summary>
        public void AddSection()
        {
            AddElementToDocument(new XElement(Name.Paragraph,
                          new XAttribute(Name.ParagraphId, HelperFunctions.GenerateHexId()),
                          new XElement(Name.ParagraphProperties,
                              new XElement(Name.SectionProperties,
                                  new XElement(Namespace.Main + "type",
                                      new XAttribute(Name.MainVal, SectionBreakType.Continuous.GetEnumName()))))));
        }

        /// <summary>
        /// Add a new page break to the container
        /// </summary>
        public void AddPageBreak()
        {
            AddElementToDocument(new XElement(Name.Paragraph,
                        new XAttribute(Name.ParagraphId, HelperFunctions.GenerateHexId()),
                        new XElement(Name.ParagraphProperties,
                            new XElement(Name.SectionProperties))));
        }

        /// <summary>
        /// Add a paragraph with the given text to the end of the container
        /// </summary>
        /// <param name="text">Text to add</param>
        /// <param name="formatting">Formatting to use</param>
        /// <returns></returns>
        public Paragraph AddParagraph(string text, Formatting formatting)
        {
            if (text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            XElement paragraph = Paragraph.Create(text, formatting);
            AddElementToDocument(paragraph);
            return OnAddParagraph(new Paragraph(Document, paragraph, 0));
        }

        /// <summary>
        /// Add a new table to the end of the container
        /// </summary>
        /// <param name="table">Table to add</param>
        /// <returns>The table now associated with the document.</returns>
        public Table AddTable(Table table)
        {
            if (table.InDom)
                throw new ArgumentException("Cannot add table multiple times.", nameof(table));

            table.BlockContainer = this;
            AddElementToDocument(table.Xml);

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
            if (table.InDom)
                throw new ArgumentException("Cannot add table multiple times.", nameof(table));

            table.BlockContainer = this;

            Paragraph firstParagraph = Document.FindParagraphByIndex(index);
            XElement[] split = SplitParagraph(firstParagraph, index - firstParagraph.StartIndex);
            firstParagraph.Xml.ReplaceWith(split[0], table.Xml, split[1]);

            return table;
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
                Paragraphs.ToList().ForEach(InsertMissingStyles);
            }
        }
        
        /// <summary>
        /// Split a paragraph at a specific index
        /// </summary>
        /// <param name="paragraph">FirstParagraph to split</param>
        /// <param name="index">Character index to split at</param>
        /// <returns>Left/Right split</returns>
        /// <exception cref="ArgumentNullException"></exception>
        private static XElement[] SplitParagraph(Paragraph paragraph, int index)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));
            
            if (index < 0)
                throw new ArgumentOutOfRangeException(nameof(index));

            Run r = paragraph.FindRunAffectedByEdit(EditType.Insert, index);
            XElement[] split;
            XElement before, after;

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
}