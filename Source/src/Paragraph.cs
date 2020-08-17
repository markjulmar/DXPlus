using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
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
        internal List<XElement> runs;
        internal List<XElement> styles;
        private Alignment alignment;            // This paragraphs text alignment
        private Direction direction;
        internal readonly int startIndex;
        internal readonly int endIndex;
        private float indentationAfter;
        private float indentationBefore;
        private float indentationFirstLine;
        private float indentationHanging;
        private int? indentLevel;
        private XElement paragraphNumberProperties;

        internal Paragraph(DocX document, XElement xml, int startIndex, ContainerType parent = ContainerType.None)
            : base(document, xml)
        {
            ParentContainer = parent;
            this.startIndex = startIndex;
            endIndex = startIndex + GetElementTextLength(Xml);
            styles = new List<XElement>();

            DocumentProperties = Xml.Descendants(DocxNamespace.Main + "fldSimple")
                                    .Select(xml => new DocProperty(Document, xml))
                                    .ToList();

            runs = Xml.Elements(DocxNamespace.Main + "r").ToList();
        }

        /// <summary>
        /// Gets or set this Paragraphs text alignment.
        /// </summary>
        public Alignment Alignment
        {
            get
            {
                XElement jc = ParaProperties().Element(DocxNamespace.Main + "jc");
                if (jc != null && jc.TryGetEnumValue<Alignment>(out Alignment result))
                {
                    return result;
                }
                return Alignment.Left;
            }

            set
            {
                alignment = value;

                XElement pPr = ParaProperties();
                XElement jc = pPr.Element(DocxNamespace.Main + "jc");

                if (alignment != Alignment.Left)
                {
                    if (jc == null)
                    {
                        pPr.Add(new XElement(DocxNamespace.Main + "jc",
                                    new XAttribute(DocxNamespace.Main + "val", alignment.GetEnumName())));
                    }
                    else
                    {
                        jc.SetAttributeValue(DocxNamespace.Main + "val", alignment.GetEnumName());
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
            get
            {
                XElement pPr = ParaProperties();
                XElement bidi = pPr.Element(DocxNamespace.Main + "bidi");
                return bidi == null ? Direction.LeftToRight : Direction.RightToLeft;
            }

            set
            {
                direction = value;

                XElement pPr = ParaProperties();
                XElement bidi = pPr.Element(DocxNamespace.Main + "bidi");

                if (direction == Direction.RightToLeft)
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
        /// Returns a list of field type DocProperty in this document.
        /// </summary>
        public List<DocProperty> DocumentProperties { get; }

        ///<summary>
        /// Returns table following the paragraph. Null if the following element isn't table.
        ///</summary>
        public Table FollowingTable { get; internal set; }

        /// <summary>
        /// Add a heading
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="headingType"></param>
        /// <returns></returns>
        public Paragraph Heading(HeadingType headingType)
        {
            StyleName = headingType.GetEnumName();
            return this;
        }

        /// <summary>
        /// Returns a list of Hyperlinks in this Paragraph.
        /// </summary>
        public List<Hyperlink> Hyperlinks
        {
            get
            {
                List<Hyperlink> hyperlinks = new List<Hyperlink>();

                foreach (XElement he in Xml.Descendants().Where(h => h.Name.LocalName == "hyperlink" || h.Name.LocalName == "instrText").ToList())
                {
                    if (he.Name.LocalName == "hyperlink")
                    {
                        hyperlinks.Add(new Hyperlink(Document, he, uri: null)
                        {
                            packagePart = packagePart
                        });
                    }
                    else
                    {
                        // Find the parent run, no matter how deeply nested we are.
                        XElement e = he;
                        while (e.Name.LocalName != "r")
                        {
                            e = e.Parent;
                        }

                        // Take every element until we reach w:fldCharType="end"
                        List<XElement> hyperlink_runs = new List<XElement>();
                        foreach (XElement r in e.ElementsAfterSelf(DocxNamespace.Main + "r"))
                        {
                            // Add this run to the list.
                            hyperlink_runs.Add(r);

                            XElement fldChar = r.Descendants(DocxNamespace.Main + "fldChar").SingleOrDefault();
                            if (fldChar != null)
                            {
                                XAttribute fldCharType = fldChar.Attribute(DocxNamespace.Main + "fldCharType");
                                if (fldCharType?.Value.Equals("end", StringComparison.CurrentCultureIgnoreCase) == true)
                                {
                                    hyperlinks.Add(new Hyperlink(Document, he, hyperlink_runs)
                                    {
                                        packagePart = packagePart
                                    });
                                    break;
                                }
                            }
                        }
                    }
                }

                return hyperlinks;
            }
        }

        /// <summary>
        /// Set the after indentation in cm for this Paragraph.
        /// </summary>
        public float IndentationAfter
        {
            get
            {
                XElement ind = GetOrCreate_pPr_ind();
                XAttribute right = ind.Attribute(DocxNamespace.Main + "right");
                return right != null ? float.Parse(right.Value) : 0.0f;
            }

            set
            {
                if (IndentationAfter != value)
                {
                    indentationAfter = value;

                    XElement ind = GetOrCreate_pPr_ind();

                    string indentation = (indentationAfter / 0.1 * 57).ToString();

                    XAttribute right = ind.Attribute(DocxNamespace.Main + "right");
                    if (right != null)
                    {
                        right.Value = indentation;
                    }
                    else
                    {
                        ind.Add(new XAttribute(DocxNamespace.Main + "right", indentation));
                    }
                }
            }
        }

        /// <summary>
        /// Set the before indentation in cm for this Paragraph.
        /// </summary>
        public float IndentationBefore
        {
            get
            {
                XElement ind = GetOrCreate_pPr_ind();
                XAttribute left = ind.Attribute(DocxNamespace.Main + "left");
                return left != null ? float.Parse(left.Value) / (57 * 10) : 0.0f;
            }

            set
            {
                if (IndentationBefore != value)
                {
                    indentationBefore = value;

                    XElement ind = GetOrCreate_pPr_ind();

                    string indentation = (indentationBefore / 0.1 * 57).ToString();

                    XAttribute left = ind.Attribute(DocxNamespace.Main + "left");
                    if (left != null)
                    {
                        left.Value = indentation;
                    }
                    else
                    {
                        ind.Add(new XAttribute(DocxNamespace.Main + "left", indentation));
                    }
                }
            }
        }

        /// <summary>
        /// Get or set the indentation of the first line of this Paragraph.
        /// </summary>
        public float IndentationFirstLine
        {
            get
            {
                XElement ind = GetOrCreate_pPr_ind();
                XAttribute firstLine = ind.Attribute(DocxNamespace.Main + "firstLine");

                return firstLine != null ? float.Parse(firstLine.Value) : 0.0f;
            }

            set
            {
                if (IndentationFirstLine != value)
                {
                    indentationFirstLine = value;
                    XElement ind = GetOrCreate_pPr_ind();

                    // Paragraph can either be firstLine or hanging (Remove hanging).
                    XAttribute hanging = ind.Attribute(DocxNamespace.Main + "hanging");
                    hanging?.Remove();

                    string indentation = ((indentationFirstLine / 0.1) * 57).ToString();
                    XAttribute firstLine = ind.Attribute(DocxNamespace.Main + "firstLine");
                    if (firstLine != null)
                    {
                        firstLine.Value = indentation;
                    }
                    else
                    {
                        ind.Add(new XAttribute(DocxNamespace.Main + "firstLine", indentation));
                    }
                }
            }
        }

        /// <summary>
        /// Get or set the indentation of all but the first line of this Paragraph.
        /// </summary>
        public float IndentationHanging
        {
            get
            {
                XElement ind = GetOrCreate_pPr_ind();
                XAttribute hanging = ind.Attribute(DocxNamespace.Main + "hanging");
                return hanging != null ? float.Parse(hanging.Value) / (57 * 10) : 0.0f;
            }

            set
            {
                if (IndentationHanging != value)
                {
                    indentationHanging = value;
                    XElement ind = GetOrCreate_pPr_ind();

                    // Paragraph can either be firstLine or hanging (Remove firstLine).
                    XAttribute firstLine = ind.Attribute(DocxNamespace.Main + "firstLine");
                    firstLine?.Remove();

                    string indentation = (indentationHanging / 0.1 * 57).ToString();
                    XAttribute hanging = ind.Attribute(DocxNamespace.Main + "hanging");
                    if (hanging != null)
                    {
                        hanging.Value = indentation;
                    }
                    else
                    {
                        ind.Add(new XAttribute(DocxNamespace.Main + "hanging", indentation));
                    }
                }
            }
        }

        /// <summary>
        /// If this element is a list item, get the indentation level of the list item.
        /// </summary>
        public int? IndentLevel => !IsListItem
                    ? null
                    : (int?)(indentLevel ??= int.Parse(ParagraphNumberProperties.FirstLocalNameDescendant("ilvl").GetVal()));

        public bool ShouldKeepWithNext
        {
            get
            {
                XElement pPr = ParaProperties();
                XElement keepWithNextE = pPr.Element(DocxNamespace.Main + "keepNext");
                return keepWithNextE != null;
            }
        }

        /// <summary>
        /// Determine if this paragraph is a list element.
        /// </summary>
        public bool IsListItem => ParagraphNumberProperties != null;

        public float LineSpacing
        {
            get
            {
                XElement pPr = ParaProperties();
                XElement spacing = pPr.Element(DocxNamespace.Main + "spacing");

                if (spacing != null)
                {
                    XAttribute line = spacing.Attribute(DocxNamespace.Main + "line");
                    if (line != null && float.TryParse(line.Value, out float f))
                    {
                        return f / 20.0f;
                    }
                }

                return 1.1f * 20.0f;
            }

            set => Spacing(value);
        }

        public float LineSpacingAfter
        {
            get
            {
                XElement pPr = ParaProperties();
                XElement spacing = pPr.Element(DocxNamespace.Main + "spacing");

                if (spacing != null)
                {
                    XAttribute line = spacing.Attribute(DocxNamespace.Main + "after");
                    if (line != null && float.TryParse(line.Value, out float f))
                    {
                        return f / 20.0f;
                    }
                }

                return 10.0f;
            }

            set => SpacingAfter(value);
        }

        public float LineSpacingBefore
        {
            get
            {
                XElement pPr = ParaProperties();
                XElement spacing = pPr.Element(DocxNamespace.Main + "spacing");

                if (spacing != null)
                {
                    XAttribute line = spacing.Attribute(DocxNamespace.Main + "before");
                    if (line != null && float.TryParse(line.Value, out float f))
                    {
                        return f / 20.0f;
                    }
                }

                return 0.0f;
            }

            set => SpacingBefore(value);
        }

        /// <summary>
        /// Determine if the list element is a numbered list of bulleted list element
        /// </summary>
        public ListItemType ListItemType { get; set; }

        /// <summary>
        /// Gets the formatted text value of this Paragraph.
        /// </summary>
        public List<FormattedText> MagicText => HelperFunctions.GetFormattedText(Xml);

        /// <summary>
        /// Fetch the paragraph number properties for a list element.
        /// </summary>
        public XElement ParagraphNumberProperties
        {
            get
            {
                if (paragraphNumberProperties == null)
                {
                    XElement node = Xml.FirstLocalNameDescendant("numPr");
                    paragraphNumberProperties = node.FirstLocalNameDescendant("numId").GetVal() == "0" ? null : node;
                }

                return paragraphNumberProperties;
            }
        }

        public ContainerType ParentContainer { get; set; }

        /// <summary>
        /// Returns a list of all Pictures in a Paragraph.
        /// </summary>
        public List<Picture> Pictures => (
                    from p in Xml.LocalNameDescendants("drawing")
                    let id = p.FirstLocalNameDescendant("blip").AttributeValue(DocxNamespace.RelatedDoc + "embed")
                    where id != null
                    let img = new Image(Document, packagePart.GetRelationship(id))
                    select new Picture(Document, p, img)
                ).Union(
                    from p in Xml.LocalNameDescendants("pict")
                    let id = p.FirstLocalNameDescendant("imagedata").AttributeValue(DocxNamespace.RelatedDoc + "id")
                    where id != null
                    let img = new Image(Document, packagePart.GetRelationship(id))
                    select new Picture(Document, p, img)
                ).ToList();

        ///<summary>
        /// The style name of the paragraph.
        ///</summary>
        public string StyleName
        {
            get
            {
                XElement element = ParaProperties();
                XElement styleElement = element.Element(DocxNamespace.Main + "pStyle");
                if (styleElement != null)
                {
                    XAttribute attr = styleElement.Attribute(DocxNamespace.Main + "val");
                    if (attr != null && !string.IsNullOrEmpty(attr.Value))
                    {
                        return attr.Value;
                    }
                }
                return "Normal";
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    value = "Normal";
                }
                XElement element = ParaProperties();
                XElement styleElement = element.Element(DocxNamespace.Main + "pStyle");
                if (styleElement == null)
                {
                    element.Add(new XElement(DocxNamespace.Main + "pStyle"));
                    styleElement = element.Element(DocxNamespace.Main + "pStyle");
                }
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
        /// <returns>This Paragraph with the new text appened.</returns>
        public Paragraph Append(string text)
        {
            List<XElement> newRuns = HelperFunctions.FormatInput(text, null);
            Xml.Add(newRuns);
            runs = Xml.Elements(DocxNamespace.Main + "r").Reverse().Take(newRuns.Count).ToList();

            return this;
        }

        public Paragraph AppendBookmark(string bookmarkName)
        {
            XElement wBookmarkStart = new XElement(
                DocxNamespace.Main + "bookmarkStart",
                new XAttribute(DocxNamespace.Main + "id", 0),
                new XAttribute(DocxNamespace.Main + "name", bookmarkName));
            Xml.Add(wBookmarkStart);

            XElement wBookmarkEnd = new XElement(
                DocxNamespace.Main + "bookmarkEnd",
                new XAttribute(DocxNamespace.Main + "id", 0),
                new XAttribute(DocxNamespace.Main + "name", bookmarkName));
            Xml.Add(wBookmarkEnd);

            return this;
        }

        /// <summary>
        /// Append a field of type document property, this field will display the custom property cp, at the end of this paragraph.
        /// </summary>
        /// <param name="cp">The custom property to display.</param>
        /// <param name="trackChanges"></param>
        /// <param name="f">The formatting to use for this text.</param>
        /// <example>
        /// Create, add and display a custom property in a document.
        /// <code>
        /// // Load a document.
        ///using (DocX document = DocX.Create("CustomProperty_Add.docx"))
        ///{
        ///    // Add a few Custom Properties to this document.
        ///    document.AddCustomProperty(new CustomProperty("fname", "cathal"));
        ///    document.AddCustomProperty(new CustomProperty("age", 24));
        ///    document.AddCustomProperty(new CustomProperty("male", true));
        ///    document.AddCustomProperty(new CustomProperty("newyear2012", new DateTime(2012, 1, 1)));
        ///    document.AddCustomProperty(new CustomProperty("fav_num", 3.141592));
        ///
        ///    // Insert a new Paragraph and append a load of DocProperties.
        ///    Paragraph p = document.InsertParagraph("fname: ")
        ///        .AppendDocProperty(document.CustomProperties["fname"])
        ///        .Append(", age: ")
        ///        .AppendDocProperty(document.CustomProperties["age"])
        ///        .Append(", male: ")
        ///        .AppendDocProperty(document.CustomProperties["male"])
        ///        .Append(", newyear2012: ")
        ///        .AppendDocProperty(document.CustomProperties["newyear2012"])
        ///        .Append(", fav_num: ")
        ///        .AppendDocProperty(document.CustomProperties["fav_num"]);
        ///
        ///    // Save the changes to the document.
        ///    document.Save();
        ///}
        /// </code>
        /// </example>
        public Paragraph AppendDocProperty(CustomProperty cp, bool trackChanges = false, Formatting f = null)
        {
            InsertDocProperty(cp, trackChanges, f);
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
            XElement oMathPara =
                new XElement(DocxNamespace.Math + "oMathPara",
                    new XElement(DocxNamespace.Math + "oMath",
                        new XElement(DocxNamespace.Main + "r",
                            new Formatting() { FontFamily = new FontFamily("Cambria Math") }.Xml,
                            new XElement(DocxNamespace.Math + "t", equation)
                        )
                    )
                );

            // Add equation element into paragraph xml and update runs collection
            Xml.Add(oMathPara);
            runs = Xml.Elements(DocxNamespace.Math + "oMathPara").ToList();

            // Return paragraph with equation
            return this;
        }

        /// <summary>
        /// Append a hyperlink to a Paragraph.
        /// </summary>
        /// <param name="h">The hyperlink to append.</param>
        /// <returns>The Paragraph with the hyperlink appended.</returns>
        public Paragraph AppendHyperlink(Hyperlink h)
        {
            // Convert the path of this mainPart to its equilivant rels file path.
            string path = packagePart.Uri.OriginalString.Replace("/word/", "");
            Uri rels_path = new Uri("/word/_rels/" + path + ".rels", UriKind.Relative);

            // Check to see if the rels file exists and create it if not.
            if (!Document.package.PartExists(rels_path))
            {
                HelperFunctions.CreateRelsPackagePart(Document, rels_path);
            }

            // Check to see if a rel for this Hyperlink exists, create it if not.
            string Id = GetOrGenerateRel(h);

            Xml.Add(h.Xml);
            Xml.Elements().Last().SetAttributeValue(DocxNamespace.RelatedDoc + "id", Id);

            runs = Xml.Elements().Last().Elements(DocxNamespace.Main + "r").ToList();

            return this;
        }

        /// <summary>
        /// Append text on a new line to this Paragraph.
        /// </summary>
        /// <param name="text">The text to append.</param>
        /// <returns>This Paragraph with the new text appened.</returns>
        /// <example>
        /// Add a new Paragraph to this document and then append a new line with some text to it.
        /// <code>
        /// // Load a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph and Append a new line with some text to it.
        ///     Paragraph p = document.InsertParagraph().AppendLine("Hello World!!!");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }
        /// </code>
        /// </example>
        public Paragraph AppendLine(string text)
        {
            return Append("\n" + text);
        }

        /// <summary>
        /// Append a new line to this Paragraph.
        /// </summary>
        /// <returns>This Paragraph with a new line appeneded.</returns>
        public Paragraph AppendLine()
        {
            return Append("\n");
        }

        /// <summary>
        /// Append a PageCount place holder onto the end of a Paragraph.
        /// </summary>
        /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        public void AppendPageCount(PageNumberFormat pnf)
        {
            XElement fldSimple = new XElement(DocxNamespace.Main + "fldSimple");

            if (pnf == PageNumberFormat.Normal)
            {
                fldSimple.Add(new XAttribute(DocxNamespace.Main + "instr", @" NUMPAGES   \* MERGEFORMAT "));
            }
            else
            {
                fldSimple.Add(new XAttribute(DocxNamespace.Main + "instr", @" NUMPAGES  \* ROMAN  \* MERGEFORMAT "));
            }

            XElement content = XElement.Parse
            (
             @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
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
        /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add Headers to the document.
        ///     document.AddHeaders();
        ///
        ///     // Get the default Header.
        ///     Header header = document.Headers.odd;
        ///
        ///     // Insert a Paragraph into the Header.
        ///     Paragraph p0 = header.InsertParagraph();
        ///
        ///     // Appemd place holders for PageNumber and PageCount into the Header.
        ///     // Word will replace these with the correct value for each Page.
        ///     p0.Append("Page (");
        ///     p0.AppendPageNumber(PageNumberFormat.normal);
        ///     p0.Append(" of ");
        ///     p0.AppendPageCount(PageNumberFormat.normal);
        ///     p0.Append(")");
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="AppendPageCount"/>
        /// <seealso cref="InsertPageNumber"/>
        /// <seealso cref="InsertPageCount"/>
        public void AppendPageNumber(PageNumberFormat pnf)
        {
            XElement fldSimple = new XElement(DocxNamespace.Main + "fldSimple");

            if (pnf == PageNumberFormat.Normal)
            {
                fldSimple.Add(new XAttribute(DocxNamespace.Main + "instr", @" PAGE   \* MERGEFORMAT "));
            }
            else
            {
                fldSimple.Add(new XAttribute(DocxNamespace.Main + "instr", @" PAGE  \* ROMAN  \* MERGEFORMAT "));
            }

            XElement content = XElement.Parse
            (
             @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
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
            // Convert the path of this mainPart to its equilivant rels file path.
            string path = packagePart.Uri.OriginalString.Replace("/word/", "");
            Uri rels_path = new Uri("/word/_rels/" + path + ".rels", UriKind.Relative);

            // Check to see if the rels file exists and create it if not.
            if (!Document.package.PartExists(rels_path))
            {
                HelperFunctions.CreateRelsPackagePart(Document, rels_path);
            }

            // Check to see if a rel for this Picture exists, create it if not.
            string Id = GetOrGenerateRel(p);

            // Add the Picture Xml to the end of the Paragragraph Xml.
            Xml.Add(p.Xml);

            // Extract the attribute id from the Pictures Xml.
            XAttribute a_id = Xml.Elements().Last().Descendants()
                            .Where(e => e.Name.LocalName.Equals("blip"))
                            .Select(e => e.Attribute(DocxNamespace.RelatedDoc + "embed"))
                            .Single();

            // Set its value to the Pictures relationships id.
            a_id.SetValue(Id);

            // For formatting such as .Bold()
            runs = Xml.Elements(DocxNamespace.Main + "r").Reverse()
                      .Take(p.Xml.Elements(DocxNamespace.Main + "r")
                      .Count()).ToList();

            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <returns>This Paragraph with the last appended text bold.</returns>
        /// <example>
        /// Append text to this Paragraph and then make it bold.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("Bold").Bold()
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph Bold()
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "b", string.Empty, null);
            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="capsStyle">The caps style to apply to the last appended text.</param>
        /// <returns>This Paragraph with the last appended text's caps style changed.</returns>
        /// <example>
        /// Append text to this Paragraph and then set it to full caps.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("Capitalized").CapsStyle(CapsStyle.caps)
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph CapsStyle(CapsStyle capsStyle)
        {
            if (capsStyle != DXPlus.CapsStyle.None)
            {
                ApplyTextFormattingProperty(DocxNamespace.Main + capsStyle.ToString(), string.Empty, null);
            }

            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="c">A color to use on the appended text.</param>
        /// <returns>This Paragraph with the last appended text colored.</returns>
        /// <example>
        /// Append text to this Paragraph and then color it.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("Blue").Color(Color.Blue)
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph Color(Color c)
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "color", string.Empty, new XAttribute(DocxNamespace.Main + "val", c.ToHex()));
            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="culture">The CultureInfo for text</param>
        /// <returns>This Paragraph in curent culture</returns>
        public Paragraph Culture(CultureInfo culture)
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "lang", string.Empty,
                new XAttribute(DocxNamespace.Main + "val", culture.Name));
            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <returns>This Paragraph in curent culture</returns>
        public Paragraph CurentCulture()
        {
            ApplyTextFormattingProperty(DocxNamespace.Main + "lang",
                string.Empty, new XAttribute(DocxNamespace.Main + "val", CultureInfo.CurrentCulture.Name));
            return this;
        }

        /// <summary>
        /// Find all instances of a string in this paragraph and return their indexes in a List.
        /// </summary>
        /// <param name="text">The string to find</param>
        /// <param name="options">The options to use when finding a string match.</param>
        /// <returns>A list of indexes.</returns>
        public IEnumerable<int> FindAll(string text, RegexOptions options)
        {
            MatchCollection mc = Regex.Matches(Text, Regex.Escape(text), options);
            return mc.Cast<Match>().Select(m => m.Index);
        }

        /// <summary>
        ///  Find all unique instances of the given Regex Pattern
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public IEnumerable<(int index, string text)> FindPattern(string pattern, RegexOptions options)
        {
            MatchCollection mc = Regex.Matches(Text, pattern, options);
            return mc.Cast<Match>().Select(m => (index: m.Index, text: m.Value));
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="fontFamily">The font to use for the appended text.</param>
        /// <returns>This Paragraph with the last appended text's font changed.</returns>
        /// <example>
        /// Append text to this Paragraph and then change its font.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("Times new roman").Font(new FontFamily("Times new roman"))
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph Font(FontFamily fontFamily)
        {
            ApplyTextFormattingProperty
            (
                DocxNamespace.Main + "rFonts",
                string.Empty,
                new[]
                {
                    new XAttribute(DocxNamespace.Main + "ascii", fontFamily.Name),
                    new XAttribute(DocxNamespace.Main + "hAnsi", fontFamily.Name), // Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
                    new XAttribute(DocxNamespace.Main + "cs", fontFamily.Name),    // Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
                }
            );

            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="fontSize">The font size to use for the appended text.</param>
        /// <returns>This Paragraph with the last appended text resized.</returns>
        /// <example>
        /// Append text to this Paragraph and then resize it.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("Big").FontSize(20)
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph FontSize(double fontSize)
        {
            if (fontSize < 0 || fontSize > 1639)
            {
                throw new ArgumentException("Size", "Value must be in the range 0 - 1638");
            }

            if (fontSize - (int)fontSize != 0)
            {
                throw new ArgumentException("Size", "Value must be either a whole or half number, examples: 32, 32.5");
            }

            ApplyTextFormattingProperty(DocxNamespace.Main + "sz", string.Empty, new XAttribute(DocxNamespace.Main + "val", fontSize * 2));
            ApplyTextFormattingProperty(DocxNamespace.Main + "szCs", string.Empty, new XAttribute(DocxNamespace.Main + "val", fontSize * 2));

            return this;
        }

        public IEnumerable<Bookmark> GetBookmarks()
        {
            return Xml.Descendants(DocxNamespace.Main + "bookmarkStart")
                        .Select(x => x.Attribute(DocxNamespace.Main + "name"))
                        .Select(x => new Bookmark(x.Value, this));
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <returns>This Paragraph with the last appended text hidden.</returns>
        /// <example>
        /// Append text to this Paragraph and then hide it.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("hidden").Hide()
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
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

        public void InsertAtBookmark(string toInsert, string bookmarkName)
        {
            XElement bookmark = Xml.Descendants(DocxNamespace.Main + "bookmarkStart")
                                .SingleOrDefault(x => x.Attribute(DocxNamespace.Main + "name").Value == bookmarkName);
            if (bookmark != null)
            {
                List<XElement> run = HelperFunctions.FormatInput(toInsert, null);
                bookmark.AddBeforeSelf(run);
                runs = Xml.Elements(DocxNamespace.Main + "r").ToList();
                HelperFunctions.RenumberIDs(Document);
            }
        }

        /// <summary>
        /// Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
        /// </summary>
        /// <param name="cp">The custom property to display.</param>
        /// <param name="trackChanges"></param>
        /// <param name="f">The formatting to use for this text.</param>
        /// <example>
        /// Create, add and display a custom property in a document.
        /// <code>
        /// // Load a document
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Create a custom property.
        ///     CustomProperty name = new CustomProperty("name", "Cathal Coffey");
        ///
        ///     // Add this custom property to this document.
        ///     document.AddCustomProperty(name);
        ///
        ///     // Create a text formatting.
        ///     Formatting f = new Formatting();
        ///     f.Bold = true;
        ///     f.Size = 14;
        ///     f.StrikeThrough = StrickThrough.strike;
        ///
        ///     // Insert a new paragraph.
        ///     Paragraph p = document.InsertParagraph("Author: ", false, f);
        ///
        ///     // Insert a field of type document property to display the custom property name and track this change.
        ///     p.InsertDocProperty(name, true, f);
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public DocProperty InsertDocProperty(CustomProperty cp, bool trackChanges = false, Formatting f = null)
        {
            XElement f_xml = null;
            if (f != null)
            {
                f_xml = f.Xml;
            }

            XElement e = new XElement
            (
                DocxNamespace.Main + "fldSimple",
                new XAttribute(DocxNamespace.Main + "instr", string.Format(@"DOCPROPERTY {0} \* MERGEFORMAT", cp.Name)),
                    new XElement(DocxNamespace.Main + "r",
                        new XElement(DocxNamespace.Main + "t", f_xml, cp.Value))
            );

            XElement xml = e;
            if (trackChanges)
            {
                DateTime now = DateTime.Now;
                DateTime insert_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);
                e = HelperFunctions.CreateEdit(EditType.Ins, insert_datetime, e);
            }

            Xml.Add(e);

            return new DocProperty(Document, xml);
        }

        /// <summary>
        /// This function inserts a hyperlink into a Paragraph at a specified character index.
        /// </summary>
        /// <param name="h">The hyperlink to insert.</param>
        /// <param name="index">The index to insert at.</param>
        /// <returns>The Paragraph with the Hyperlink inserted at the specified index.</returns>
        public Paragraph InsertHyperlink(Hyperlink h, int index = 0)
        {
            // Convert the path of this mainPart to its equilivant rels file path.
            string path = packagePart.Uri.OriginalString.Replace("/word/", "");
            Uri rels_path = new Uri(string.Format("/word/_rels/{0}.rels", path), UriKind.Relative);

            // Check to see if the rels file exists and create it if not.
            if (!Document.package.PartExists(rels_path))
            {
                HelperFunctions.CreateRelsPackagePart(Document, rels_path);
            }

            // Check to see if a rel for this Picture exists, create it if not.
            string Id = GetOrGenerateRel(h);

            XElement h_xml;
            if (index == 0)
            {
                // Add this hyperlink as the last element.
                Xml.AddFirst(h.Xml);

                // Extract the picture back out of the DOM.
                h_xml = (XElement)Xml.FirstNode;
            }
            else
            {
                // Get the first run effected by this Insert
                Run run = GetFirstRunEffectedByEdit(index);

                if (run == null)
                {
                    // Add this hyperlink as the last element.
                    Xml.Add(h.Xml);

                    // Extract the picture back out of the DOM.
                    h_xml = (XElement)Xml.LastNode;
                }
                else
                {
                    // Split this run at the point you want to insert
                    XElement[] splitRun = Run.SplitRun(run, index);

                    // Replace the origional run.
                    run.Xml.ReplaceWith
                    (
                        splitRun[0],
                        h.Xml,
                        splitRun[1]
                    );

                    // Get the first run effected by this Insert
                    run = GetFirstRunEffectedByEdit(index);

                    // The picture has to be the next element, extract it back out of the DOM.
                    h_xml = (XElement)run.Xml.NextNode;
                }
            }

            h_xml.SetAttributeValue(DocxNamespace.RelatedDoc + "id", Id);

            return this;
        }

        /// <summary>
        /// Insert a PageCount place holder into a Paragraph.
        /// This place holder should only be inserted into a Header or Footer Paragraph.
        /// Word will not automatically update this field if it is inserted into a document level Paragraph.
        /// </summary>
        /// <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
        /// <param name="index">The text index to insert this PageCount place holder at.</param>
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add Headers to the document.
        ///     document.AddHeaders();
        ///
        ///     // Get the default Header.
        ///     Header header = document.Headers.odd;
        ///
        ///     // Insert a Paragraph into the Header.
        ///     Paragraph p0 = header.InsertParagraph("Page ( of )");
        ///
        ///     // Insert place holders for PageNumber and PageCount into the Header.
        ///     // Word will replace these with the correct value for each Page.
        ///     p0.InsertPageNumber(PageNumberFormat.normal, 6);
        ///     p0.InsertPageCount(PageNumberFormat.normal, 11);
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="AppendPageCount"/>
        /// <seealso cref="AppendPageNumber"/>
        /// <seealso cref="InsertPageNumber"/>
        public void InsertPageCount(PageNumberFormat pnf, int index = 0)
        {
            XElement fldSimple = new XElement(DocxNamespace.Main + "fldSimple");

            if (pnf == PageNumberFormat.Normal)
            {
                fldSimple.Add(new XAttribute(DocxNamespace.Main + "instr", @" NUMPAGES   \* MERGEFORMAT "));
            }
            else
            {
                fldSimple.Add(new XAttribute(DocxNamespace.Main + "instr", @" NUMPAGES  \* ROMAN  \* MERGEFORMAT "));
            }

            XElement content = XElement.Parse(
             @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
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
                Run r = GetFirstRunEffectedByEdit(index, EditType.Ins);
                XElement[] splitEdit = SplitEdit(r.Xml, index, EditType.Ins);
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
        /// <example>
        /// <code>
        /// // Create a new document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Add Headers to the document.
        ///     document.AddHeaders();
        ///
        ///     // Get the default Header.
        ///     Header header = document.Headers.odd;
        ///
        ///     // Insert a Paragraph into the Header.
        ///     Paragraph p0 = header.InsertParagraph("Page ( of )");
        ///
        ///     // Insert place holders for PageNumber and PageCount into the Header.
        ///     // Word will replace these with the correct value for each Page.
        ///     p0.InsertPageNumber(PageNumberFormat.normal, 6);
        ///     p0.InsertPageCount(PageNumberFormat.normal, 11);
        ///
        ///     // Save the document.
        ///     document.Save();
        /// }
        /// </code>
        /// </example>
        /// <seealso cref="AppendPageCount"/>
        /// <seealso cref="AppendPageNumber"/>
        /// <seealso cref="InsertPageCount"/>
        public void InsertPageNumber(PageNumberFormat pnf, int index = 0)
        {
            XElement fldSimple = new XElement(DocxNamespace.Main + "fldSimple");

            if (pnf == PageNumberFormat.Normal)
            {
                fldSimple.Add(new XAttribute(DocxNamespace.Main + "instr", @" PAGE   \* MERGEFORMAT "));
            }
            else
            {
                fldSimple.Add(new XAttribute(DocxNamespace.Main + "instr", @" PAGE  \* ROMAN  \* MERGEFORMAT "));
            }

            XElement content = XElement.Parse(
             @"<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
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
                Run r = GetFirstRunEffectedByEdit(index, EditType.Ins);
                XElement[] splitEdit = SplitEdit(r.Xml, index, EditType.Ins);
                r.Xml.ReplaceWith(splitEdit[0], fldSimple, splitEdit[1]);
            }
        }

        /// <summary>
        /// Insert a Picture into a Paragraph at the given text index.
        /// If not index is provided defaults to 0.
        /// </summary>
        /// <param name="p">The Picture to insert.</param>
        /// <param name="index">The text index to insert at.</param>
        /// <returns>The modified Paragraph.</returns>
        public Paragraph InsertPicture(Picture p, int index = 0)
        {
            // Convert the path of this mainPart to its equilivant rels file path.
            string path = packagePart.Uri.OriginalString.Replace("/word/", "");
            Uri rels_path = new Uri("/word/_rels/" + path + ".rels", UriKind.Relative);

            // Check to see if the rels file exists and create it if not.
            if (!Document.package.PartExists(rels_path))
            {
                HelperFunctions.CreateRelsPackagePart(Document, rels_path);
            }

            // Check to see if a rel for this Picture exists, create it if not.
            string Id = GetOrGenerateRel(p);

            XElement p_xml;
            if (index == 0)
            {
                // Add this hyperlink as the last element.
                Xml.AddFirst(p.Xml);

                // Extract the picture back out of the DOM.
                p_xml = (XElement)Xml.FirstNode;
            }
            else
            {
                // Get the first run effected by this Insert
                Run run = GetFirstRunEffectedByEdit(index);

                if (run == null)
                {
                    // Add this picture as the last element.
                    Xml.Add(p.Xml);

                    // Extract the picture back out of the DOM.
                    p_xml = (XElement)Xml.LastNode;
                }
                else
                {
                    // Split this run at the point you want to insert
                    XElement[] splitRun = Run.SplitRun(run, index);

                    // Replace the origional run.
                    run.Xml.ReplaceWith
                    (
                        splitRun[0],
                        p.Xml,
                        splitRun[1]
                    );

                    // Get the first run effected by this Insert
                    run = GetFirstRunEffectedByEdit(index);

                    // The picture has to be the next element, extract it back out of the DOM.
                    p_xml = (XElement)run.Xml.NextNode;
                }
            }
            // Extract the attribute id from the Pictures Xml.
            XAttribute a_id = p_xml.LocalNameDescendants("blip")
                    .Select(e => e.Attribute(DocxNamespace.RelatedDoc + "embed"))
                    .Single();

            // Set its value to the Pictures relationships id.
            a_id.SetValue(Id);

            return this;
        }

        /// <summary>
        /// Inserts a sstring into a Paragraph at a specified index position.
        /// </summary>
        public void InsertText(string value, Formatting formatting = null)
        {
            if (formatting == null)
            {
                formatting = new Formatting();
            }

            List<XElement> newRuns = HelperFunctions.FormatInput(value, formatting.Xml);
            Xml.Add(newRuns);

            HelperFunctions.RenumberIDs(Document);
        }

        /// <summary>
        /// Inserts a string into a Paragraph at a specified index position.
        /// </summary>
        /// <param name="index">The index position of the insertion.</param>
        /// <param name="value">The System.String to insert.</param>
        /// <param name="trackChanges">Flag this insert as a change.</param>
        /// <param name="formatting">The text formatting.</param>
        public void InsertText(int index, string value, bool trackChanges = false, Formatting formatting = null)
        {
            // Timestamp to mark the start of insert
            DateTime now = DateTime.Now;
            DateTime insert_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

            // Get the first run effected by this Insert
            Run run = GetFirstRunEffectedByEdit(index);
            if (run == null)
            {
                object insert;
                if (formatting != null) //not sure how to get original formatting here when run == null
                {
                    insert = HelperFunctions.FormatInput(value, formatting.Xml);
                }
                else
                {
                    insert = HelperFunctions.FormatInput(value, null);
                }

                if (trackChanges)
                {
                    insert = HelperFunctions.CreateEdit(EditType.Ins, insert_datetime, insert);
                }

                Xml.Add(insert);
            }
            else
            {
                object newRuns;
                XElement rprel = run.Xml.Element(DocxNamespace.Main + "rPr");
                if (formatting != null)
                {
                    Formatting oldfmt = null;
                    if (rprel != null)
                    {
                        oldfmt = Formatting.Parse(rprel);
                    }

                    Formatting finfmt;
                    if (oldfmt != null)
                    {
                        finfmt = oldfmt.Clone();
                        if (formatting.Bold.HasValue) { finfmt.Bold = formatting.Bold; }
                        if (formatting.CapsStyle.HasValue) { finfmt.CapsStyle = formatting.CapsStyle; }
                        if (formatting.FontColor.HasValue) { finfmt.FontColor = formatting.FontColor; }
                        finfmt.FontFamily = formatting.FontFamily;
                        if (formatting.Hidden.HasValue) { finfmt.Hidden = formatting.Hidden; }
                        if (formatting.Highlight.HasValue) { finfmt.Highlight = formatting.Highlight; }
                        if (formatting.Italic.HasValue) { finfmt.Italic = formatting.Italic; }
                        if (formatting.Kerning.HasValue) { finfmt.Kerning = formatting.Kerning; }
                        finfmt.Language = formatting.Language;
                        if (formatting.Misc.HasValue) { finfmt.Misc = formatting.Misc; }
                        if (formatting.PercentageScale.HasValue) { finfmt.PercentageScale = formatting.PercentageScale; }
                        if (formatting.Position.HasValue) { finfmt.Position = formatting.Position; }
                        if (formatting.Script.HasValue) { finfmt.Script = formatting.Script; }
                        if (formatting.Size.HasValue) { finfmt.Size = formatting.Size; }
                        if (formatting.Spacing.HasValue) { finfmt.Spacing = formatting.Spacing; }
                        if (formatting.StrikeThrough.HasValue) { finfmt.StrikeThrough = formatting.StrikeThrough; }
                        if (formatting.UnderlineColor.HasValue) { finfmt.UnderlineColor = formatting.UnderlineColor; }
                        if (formatting.UnderlineStyle.HasValue) { finfmt.UnderlineStyle = formatting.UnderlineStyle; }
                    }
                    else
                    {
                        finfmt = formatting;
                    }

                    newRuns = HelperFunctions.FormatInput(value, finfmt.Xml);
                }
                else
                {
                    newRuns = HelperFunctions.FormatInput(value, rprel);
                }

                // The parent of this Run
                XElement parentElement = run.Xml.Parent;
                object insert = newRuns;

                switch (parentElement.Name.LocalName)
                {
                    case "ins":
                        // The datetime that this ins was created
                        DateTime parent_ins_date = DateTime.Parse(parentElement.Attribute(DocxNamespace.Main + "date").Value);
                        // Special case: You want to track changes, and the first Run effected by this insert
                        // has a datetime stamp equal to now.
                        if (trackChanges && parent_ins_date.CompareTo(insert_datetime) == 0)
                        {
                            // Inserting into a non edit and this special case, is the same procedure.
                            goto default;
                        }
                        goto case "del";

                    case "del":
                        if (trackChanges)
                        {
                            insert = HelperFunctions.CreateEdit(EditType.Ins, insert_datetime, newRuns);
                        }

                        // Split this Edit at the point you want to insert
                        XElement[] splitEdit = SplitEdit(parentElement, index, EditType.Ins);

                        // Replace the origional run
                        parentElement.ReplaceWith(splitEdit[0], insert, splitEdit[1]);
                        break;

                    default:
                        if (trackChanges && !parentElement.Name.LocalName.Equals("ins"))
                        {
                            _ = HelperFunctions.CreateEdit(EditType.Ins, insert_datetime, newRuns);
                        }
                        else
                        {
                            // Split this run at the point you want to insert
                            XElement[] splitRun = Run.SplitRun(run, index);

                            // Replace the origional run
                            run.Xml.ReplaceWith(splitRun[0], insert, splitRun[1]);
                        }
                        break;
                }
            }

            HelperFunctions.RenumberIDs(Document);
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <returns>This Paragraph with the last appended text italic.</returns>
        /// <example>
        /// Append text to this Paragraph and then make it italic.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("Italic").Italic()
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
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
            XElement pPr = ParaProperties();
            XElement keepLinesE = pPr.Element(DocxNamespace.Main + "keepLines");
            if (keepLinesE == null && keepTogether)
            {
                pPr.Add(new XElement(DocxNamespace.Main + "keepLines"));
            }
            if (!keepTogether && keepLinesE != null)
            {
                keepLinesE.Remove();
            }
            return this;
        }

        /// <summary>
        /// This paragraph will be kept on the same page as the next paragraph
        /// </summary>
        /// <param name="keepWithNext"></param>
        public Paragraph KeepWithNext(bool keepWithNext = true)
        {
            XElement pPr = ParaProperties();
            XElement keepWithNextE = pPr.Element(DocxNamespace.Main + "keepNext");
            if (keepWithNextE == null && keepWithNext)
            {
                pPr.Add(new XElement(DocxNamespace.Main + "keepNext"));
            }
            if (!keepWithNext && keepWithNextE != null)
            {
                keepWithNextE.Remove();
            }
            return this;
        }

        public Paragraph Kerning(int kerning)
        {
            if (!new int?[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 }.Contains(kerning))
            {
                throw new ArgumentOutOfRangeException("Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72");
            }

            ApplyTextFormattingProperty(DocxNamespace.Main + "kern", string.Empty, new XAttribute(DocxNamespace.Main + "val", kerning * 2));
            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
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

                case DXPlus.Misc.Engrave:
                    ApplyTextFormattingProperty(DocxNamespace.Main + "imprint", string.Empty, null);
                    break;

                default:
                    ApplyTextFormattingProperty(DocxNamespace.Main + misc.ToString(), string.Empty, null);
                    break;
            }

            return this;
        }

        public Paragraph PercentageScale(int percentageScale)
        {
            if (!(new int?[] { 200, 150, 100, 90, 80, 66, 50, 33 }).Contains(percentageScale))
            {
                throw new ArgumentOutOfRangeException("PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33");
            }

            ApplyTextFormattingProperty(DocxNamespace.Main + "w", string.Empty, new XAttribute(DocxNamespace.Main + "val", percentageScale));

            return this;
        }

        public Paragraph Position(double position)
        {
            if (!(position > -1585 && position < 1585))
            {
                throw new ArgumentOutOfRangeException("Position", "Value must be in the range -1585 - 1585");
            }

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
                DateTime now = DateTime.Now.ToUniversalTime();

                List<XElement> elements = Xml.Elements().ToList();
                List<XElement> temp = new List<XElement>();
                for (int i = 0; i < elements.Count; i++)
                {
                    XElement e = elements[i];

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
                if (Xml.Parent.Name.LocalName == "tc" && Xml.Parent.Elements(DocxNamespace.Main + "p").Count() == 1)
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
            // Dosen't make sense to remove a Hyperlink at a negative index.
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            // Need somewhere to store the count.
            int count = 0;
            bool found = false;
            RemoveHyperlinkRecursive(Xml, index, ref count, ref found);

            // If !found then the user tried to remove a hyperlink at an index greater than the last.
            if (!found)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
        }

        /// <summary>
        /// Removes characters from a DXPlus.DocX.Paragraph.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the paragraphs
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Remove the first two characters from every paragraph
        ///         p.RemoveText(0, 2, false);
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.InsertText(string, Formatting)"/>
        /// <param name="index">The position to begin deleting characters.</param>
        /// <param name="count">The number of characters to delete</param>
        /// <param name="trackChanges">Track changes</param>
        public void RemoveText(int index, int count, bool trackChanges = false)
        {
            // Timestamp to mark the start of insert
            DateTime now = DateTime.Now;
            DateTime remove_datetime = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);

            // The number of characters processed so far
            int processed = 0;

            do
            {
                // Get the first run effected by this Remove
                Run run = GetFirstRunEffectedByEdit(index, EditType.Del);

                // The parent of this Run
                XElement parentElement = run.Xml.Parent;
                switch (parentElement.Name.LocalName)
                {
                    case "ins":
                        {
                            XElement[] splitEditBefore = SplitEdit(parentElement, index, EditType.Del);
                            int min = Math.Min(count - processed, run.Xml.ElementsAfterSelf().Sum(e => GetElementTextLength(e)));
                            XElement[] splitEditAfter = SplitEdit(parentElement, index + min, EditType.Del);

                            XElement temp = SplitEdit(splitEditBefore[1], index + min, EditType.Del)[0];
                            object middle = HelperFunctions.CreateEdit(EditType.Del, remove_datetime, temp.Elements());
                            processed += GetElementTextLength(middle as XElement);

                            if (!trackChanges)
                            {
                                middle = null;
                            }

                            parentElement.ReplaceWith
                            (
                                splitEditBefore[0],
                                middle,
                                splitEditAfter[1]
                            );

                            processed += GetElementTextLength(middle as XElement);
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
                            XElement[] splitRunBefore = Run.SplitRun(run, index, EditType.Del);
                            int min = Math.Min(index + (count - processed), run.EndIndex);
                            XElement[] splitRunAfter = Run.SplitRun(run, min, EditType.Del);

                            object middle = HelperFunctions.CreateEdit(EditType.Del, remove_datetime, new List<XElement>()
                            {
                                Run.SplitRun(new Run(Document, splitRunBefore[1], run.StartIndex + GetElementTextLength(splitRunBefore[0])), min, EditType.Del)[0]
                            });
                            processed += GetElementTextLength(middle as XElement);

                            if (!trackChanges)
                            {
                                middle = null;
                            }

                            run.Xml.ReplaceWith(splitRunBefore[0], middle, splitRunAfter[1]);
                        }
                        break;
                }

                // If after this remove the parent element is empty, remove it.
                if (GetElementTextLength(parentElement) == 0)
                {
                    if (parentElement.Parent != null && parentElement.Parent.Name.LocalName != "tc")
                    {
                        // Need to make sure there is no drawing element within the parent element.
                        // Picture elements contain no text length but they are still content.
                        if (!parentElement.Descendants(DocxNamespace.Main + "drawing").Any())
                        {
                            parentElement.Remove();
                        }
                    }
                }
            }
            while (processed < count);

            HelperFunctions.RenumberIDs(Document);
        }

        /// <summary>
        /// Removes characters from a DXPlus.DocX.Paragraph.
        /// </summary>
        /// <example>
        /// <code>
        /// // Create a document using a relative filename.
        /// using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
        /// {
        ///     // Iterate through the paragraphs
        ///     foreach (Paragraph p in document.Paragraphs)
        ///     {
        ///         // Remove all but the first 2 characters from this Paragraph.
        ///         p.RemoveText(2, false);
        ///     }
        ///
        ///     // Save all changes made to this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        /// <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
        /// <seealso cref="Paragraph.InsertText(string, Formatting)"/>
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
            MatchCollection mc = Regex.Matches(tText, escapeRegEx ? Regex.Escape(oldValue) : oldValue, options);

            // Loop through the matches in reverse order
            foreach (Match m in mc.Cast<Match>().Reverse())
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
                        Run run = GetFirstRunEffectedByEdit(m.Index + processed);

                        // Get this runs properties
                        XElement rPr = run.Xml.Element(DocxNamespace.Main + "rPr");

                        if (rPr == null)
                        {
                            rPr = new Formatting().Xml;
                        }

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

                    //perform RegEx substitutions. Only named groups are not supported. Everything else is supported. However character escapes are not covered.
                    if (useRegExSubstitutions && !string.IsNullOrEmpty(repl))
                    {
                        repl = repl.Replace("$&", m.Value);
                        if (m.Groups.Count > 0)
                        {
                            int lastcap = 0;
                            for (int k = 0; k < m.Groups.Count; k++)
                            {
                                Group g = m.Groups[k];
                                if ((g == null) || (g.Value == ""))
                                {
                                    continue;
                                }

                                repl = repl.Replace("$" + k.ToString(), g.Value);
                                lastcap = k;
                                //cannot get named groups ATM
                            }
                            repl = repl.Replace("$+", m.Groups[lastcap].Value);
                        }
                        if (m.Index > 0)
                        {
                            repl = repl.Replace("$`", tText.Substring(0, m.Index));
                        }
                        if ((m.Index + m.Length) < tText.Length)
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
            MatchCollection matchCollection = Regex.Matches(Text, findPattern, options);

            // Loop through the matches in reverse order
            foreach (Match match in matchCollection.Cast<Match>().Reverse())
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
                        Run run = GetFirstRunEffectedByEdit(match.Index + processed);

                        // Get this runs properties
                        XElement rPr = run.Xml.Element(DocxNamespace.Main + "rPr");

                        if (rPr == null)
                        {
                            rPr = new Formatting().Xml;
                        }

                        /*
                         * Make sure that every formatting element in f.xml is also in this run,
                         * if this is not true, then their formatting does not match.
                         */
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
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="script">The script style to apply to the last appended text.</param>
        /// <returns>This Paragraph with the last appended text's script style changed.</returns>
        /// <example>
        /// Append text to this Paragraph and then set it to superscript.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("superscript").Script(Script.superscript)
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph Script(Script script)
        {
            switch (script)
            {
                case DXPlus.Script.None:
                    break;

                default:
                    {
                        ApplyTextFormattingProperty(DocxNamespace.Main + "vertAlign", string.Empty, new XAttribute(DocxNamespace.Main + "val", script.ToString()));
                        break;
                    }
            }

            return this;
        }

        /// <summary>
        /// Set the linespacing for this paragraph manually.
        /// </summary>
        /// <param name="spacingType">The type of spacing to be set, can be either Before, After or Line (Standard line spacing).</param>
        /// <param name="spacingFloat">A float value of the amount of spacing. Equals the value that van be set in Word using the "Line and Paragraph spacing" button.</param>
        public void SetLineSpacing(LineSpacingType spacingType, float spacingFloat)
        {
            spacingFloat = spacingFloat * 240;
            int spacingValue = (int)spacingFloat;

            XElement pPr = ParaProperties();
            XElement spacing = pPr.Element(DocxNamespace.Main + "spacing");
            if (spacing == null)
            {
                pPr.Add(new XElement(DocxNamespace.Main + "spacing"));
                spacing = pPr.Element(DocxNamespace.Main + "spacing");
            }

            string spacingTypeAttribute = "";
            switch (spacingType)
            {
                case LineSpacingType.Line:
                    {
                        spacingTypeAttribute = "line";
                        break;
                    }
                case LineSpacingType.Before:
                    {
                        spacingTypeAttribute = "before";
                        break;
                    }
                case LineSpacingType.After:
                    {
                        spacingTypeAttribute = "after";
                        break;
                    }
            }

            spacing.SetAttributeValue(DocxNamespace.Main + spacingTypeAttribute, spacingValue);
        }

        /// <summary>
        /// Set the linespacing for this paragraph using the Auto value.
        /// </summary>
        /// <param name="spacingType">The type of spacing to be set automatically. Using Auto will set both Before and After. None will remove any linespacing.</param>
        public void SetLineSpacing(LineSpacingTypeAuto spacingType)
        {
            int spacingValue = 100;

            XElement pPr = ParaProperties();
            XElement spacing = pPr.Element(DocxNamespace.Main + "spacing");

            if (spacingType.Equals(LineSpacingTypeAuto.None))
            {
                if (spacing != null)
                {
                    spacing.Remove();
                }
            }
            else
            {
                if (spacing == null)
                {
                    pPr.Add(new XElement(DocxNamespace.Main + "spacing"));
                    spacing = pPr.Element(DocxNamespace.Main + "spacing");
                }

                string spacingTypeAttribute = "";
                string autoSpacingTypeAttribute = "";
                switch (spacingType)
                {
                    case LineSpacingTypeAuto.AutoBefore:
                        {
                            spacingTypeAttribute = "before";
                            autoSpacingTypeAttribute = "beforeAutospacing";
                            break;
                        }
                    case LineSpacingTypeAuto.AutoAfter:
                        {
                            spacingTypeAttribute = "after";
                            autoSpacingTypeAttribute = "afterAutospacing";
                            break;
                        }
                    case LineSpacingTypeAuto.Auto:
                        {
                            spacingTypeAttribute = "before";
                            autoSpacingTypeAttribute = "beforeAutospacing";
                            spacing.SetAttributeValue(DocxNamespace.Main + "after", spacingValue);
                            spacing.SetAttributeValue(DocxNamespace.Main + "afterAutospacing", 1);
                            break;
                        }
                }

                spacing.SetAttributeValue(DocxNamespace.Main + autoSpacingTypeAttribute, 1);
                spacing.SetAttributeValue(DocxNamespace.Main + spacingTypeAttribute, spacingValue);
            }
        }

        public Paragraph Spacing(double spacing)
        {
            spacing *= 20;

            if (spacing - (int)spacing == 0)
            {
                if (!(spacing > -1585 && spacing < 1585))
                {
                    throw new ArgumentException("Spacing", "Value must be in the range: -1584 - 1584");
                }
            }
            else
            {
                throw new ArgumentException("Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9");
            }

            ApplyTextFormattingProperty(DocxNamespace.Main + "spacing", string.Empty, new XAttribute(DocxNamespace.Main + "val", spacing));

            return this;
        }

        public Paragraph SpacingAfter(double spacingAfter)
        {
            spacingAfter *= 20;

            XElement pPr = ParaProperties();
            XElement spacing = pPr.Element(DocxNamespace.Main + "spacing");
            if (spacingAfter > 0)
            {
                if (spacing == null)
                {
                    spacing = new XElement(DocxNamespace.Main + "spacing");
                    pPr.Add(spacing);
                }
                XAttribute attr = spacing.Attribute(DocxNamespace.Main + "after");
                if (attr == null)
                {
                    spacing.SetAttributeValue(DocxNamespace.Main + "after", spacingAfter);
                }
                else
                {
                    attr.SetValue(spacingAfter);
                }
            }
            if (Math.Abs(spacingAfter) < 0.1f && spacing != null)
            {
                XAttribute attr = spacing.Attribute(DocxNamespace.Main + "after");
                attr.Remove();
                if (!spacing.HasAttributes)
                {
                    spacing.Remove();
                }
            }
            //ApplyTextFormattingProperty(DocxNamespace.Main + "after", DocxNamespace.w.NamespaceName), string.Empty, new XAttribute(DocxNamespace.Main + "val", DocxNamespace.w.NamespaceName), spacingAfter));

            return this;
        }

        public Paragraph SpacingBefore(double spacingBefore)
        {
            spacingBefore *= 20;

            XElement pPr = ParaProperties();
            XElement spacing = pPr.Element(DocxNamespace.Main + "spacing");
            if (spacingBefore > 0)
            {
                if (spacing == null)
                {
                    spacing = new XElement(DocxNamespace.Main + "spacing");
                    pPr.Add(spacing);
                }
                XAttribute attr = spacing.Attribute(DocxNamespace.Main + "before");
                if (attr == null)
                {
                    spacing.SetAttributeValue(DocxNamespace.Main + "before", spacingBefore);
                }
                else
                {
                    attr.SetValue(spacingBefore);
                }
            }
            if (Math.Abs(spacingBefore) < 0.1f && spacing != null)
            {
                XAttribute attr = spacing.Attribute(DocxNamespace.Main + "before");
                attr.Remove();
                if (!spacing.HasAttributes)
                {
                    spacing.Remove();
                }
            }

            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="strikeThrough">The strike through style to used on the last appended text.</param>
        /// <returns>This Paragraph with the last appended text striked.</returns>
        public Paragraph StrikeThrough(StrikeThrough strikeThrough)
        {
            string value = strikeThrough.GetEnumName();
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
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="underlineColor">The underline color to use, if no underline is set, a single line will be used.</param>
        /// <returns>This Paragraph with the last appended text underlined in a color.</returns>
        /// <example>
        /// Append text to this Paragraph and then underline it using a color.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("color underlined").UnderlineStyle(UnderlineStyle.dotted).UnderlineColor(Color.Orange)
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph UnderlineColor(Color underlineColor)
        {
            foreach (XElement run in runs)
            {
                XElement rPr = run.Element(DocxNamespace.Main + "rPr");
                if (rPr == null)
                {
                    run.AddFirst(new XElement(DocxNamespace.Main + "rPr"));
                    rPr = run.Element(DocxNamespace.Main + "rPr");
                }

                XElement u = rPr.Element(DocxNamespace.Main + "u");
                if (u == null)
                {
                    rPr.SetElementValue(DocxNamespace.Main + "u", string.Empty);
                    u = rPr.Element(DocxNamespace.Main + "u");
                    u.SetAttributeValue(DocxNamespace.Main + "val", "single");
                }

                u.SetAttributeValue(DocxNamespace.Main + "color", underlineColor.ToHex());
            }

            return this;
        }

        /// <summary>
        /// For use with Append() and AppendLine()
        /// </summary>
        /// <param name="underlineStyle">The underline style to use for the appended text.</param>
        /// <returns>This Paragraph with the last appended text underlined.</returns>
        /// <example>
        /// Append text to this Paragraph and then underline it.
        /// <code>
        /// // Create a document.
        /// using (DocX document = DocX.Create(@"Test.docx"))
        /// {
        ///     // Insert a new Paragraph.
        ///     Paragraph p = document.InsertParagraph();
        ///
        ///     p.Append("I am ")
        ///     .Append("Underlined").UnderlineStyle(UnderlineStyle.doubleLine)
        ///     .Append(" I am not");
        ///
        ///     // Save this document.
        ///     document.Save();
        /// }// Release this document from memory.
        /// </code>
        /// </example>
        public Paragraph UnderlineStyle(UnderlineStyle underlineStyle)
        {
            string value = underlineStyle.GetEnumName();
            ApplyTextFormattingProperty(DocxNamespace.Main + "u", string.Empty, new XAttribute(DocxNamespace.Main + "val", value));
            return this;
        }

        public bool ValidateBookmark(string bookmarkName)
        {
            return GetBookmarks().Any(b => b.Name.Equals(bookmarkName));
        }

        /// <summary>
        /// Create a new Picture.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="id">A unique id that identifies an Image embedded in this document.</param>
        /// <param name="name">The name of this Picture.</param>
        /// <param name="descr">The description of this Picture.</param>
        internal static Picture CreatePicture(DocX document, string id, string name, string descr)
        {
            PackagePart part = document.package.GetPart(document.packagePart.GetRelationship(id).TargetUri);

            int newDocPrId = 1;
            List<string> existingIds = new List<string>();
            foreach (XElement bookmarkId in document.Xml.Descendants(DocxNamespace.Main + "bookmarkStart"))
            {
                XAttribute idAtt = bookmarkId.Attributes().FirstOrDefault(x => x.Name.LocalName == "id");
                if (idAtt != null)
                {
                    existingIds.Add(idAtt.Value);
                }
            }

            while (existingIds.Contains(newDocPrId.ToString()))
            {
                newDocPrId++;
            }

            int cx, cy;

            using (Stream partStream = part.GetStream())
            using (System.Drawing.Image img = System.Drawing.Image.FromStream(partStream))
            {
                cx = img.Width * 9526;
                cy = img.Height * 9526;
            }

            XElement e = new XElement(DocxNamespace.Main + "drawing");

            XElement xml = XElement.Parse
                 (string.Format(@"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:drawing xmlns = ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                        <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">
                            <wp:extent cx=""{0}"" cy=""{1}"" />
                            <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0"" />
                            <wp:docPr id=""{5}"" name=""{3}"" descr=""{4}"" />
                            <wp:cNvGraphicFramePr>
                                <a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1"" />
                            </wp:cNvGraphicFramePr>
                            <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
                                <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                    <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">
                                        <pic:nvPicPr>
                                        <pic:cNvPr id=""0"" name=""{3}"" />
                                            <pic:cNvPicPr />
                                        </pic:nvPicPr>
                                        <pic:blipFill>
                                            <a:blip r:embed=""{2}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>
                                            <a:stretch>
                                                <a:fillRect />
                                            </a:stretch>
                                        </pic:blipFill>
                                        <pic:spPr>
                                            <a:xfrm>
                                                <a:off x=""0"" y=""0"" />
                                                <a:ext cx=""{0}"" cy=""{1}"" />
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
                    ", cx, cy, id, name, descr, newDocPrId.ToString()));

            return new Picture(document, xml, new Image(document, document.packagePart.GetRelationship(id)));
        }

        internal static int GetElementTextLength(XElement run)
        {
            int count = 0;
            if (run == null)
            {
                return count;
            }

            foreach (XElement d in run.Descendants())
            {
                switch (d.Name.LocalName)
                {
                    case "tab":
                        if (d.Parent.Name.LocalName != "tabs")
                        {
                            goto case "br";
                        }

                        break;

                    case "br":
                        count++;
                        break;

                    case "t":
                    case "delText":
                        count += d.Value.Length;
                        break;
                }
            }
            return count;
        }

        internal void ApplyTextFormattingProperty(XName textFormatPropName, string value, object content)
        {
            XElement rPr;
            if (runs.Count == 0)
            {
                XElement pPr = Xml.Element(DocxNamespace.Main + "pPr");
                if (pPr == null)
                {
                    Xml.AddFirst(new XElement(DocxNamespace.Main + "pPr"));
                    pPr = Xml.Element(DocxNamespace.Main + "pPr");
                }

                rPr = pPr.Element(DocxNamespace.Main + "rPr");
                if (rPr == null)
                {
                    pPr.AddFirst(new XElement(DocxNamespace.Main + "rPr"));
                    rPr = pPr.Element(DocxNamespace.Main + "rPr");
                }

                rPr.SetElementValue(textFormatPropName, value);
                XElement last = rPr.Elements(textFormatPropName).Last();
                if (content is XAttribute attribute)
                {
                    if (last.Attribute(attribute.Name) == null)
                    {
                        last.Add(content);
                    }
                    else
                    {
                        last.Attribute(attribute.Name).Value = attribute.Value;
                    }
                }
                return;
            }

            IEnumerable<object> properties = content as IEnumerable<object>;
            bool isListOfAttributes = properties?.All(o => o is XAttribute) == true;

            foreach (XElement run in runs)
            {
                rPr = run.Element(DocxNamespace.Main + "rPr");
                if (rPr == null)
                {
                    run.AddFirst(new XElement(DocxNamespace.Main + "rPr"));
                    rPr = run.Element(DocxNamespace.Main + "rPr");
                }

                rPr.SetElementValue(textFormatPropName, value);
                XElement last = rPr.Elements(textFormatPropName).Last();

                if (isListOfAttributes)
                {
                    // List of attributes, as in the case when specifying a font family
                    foreach (XAttribute property in properties)
                    {
                        XAttribute lastAttribute = last.Attribute(property.Name);
                        if (lastAttribute == null)
                        {
                            last.Add(property);
                        }
                        else
                        {
                            lastAttribute.Value = property.Value;
                        }
                    }
                }
                else if (content is XAttribute attribute)
                {
                    if (last.Attribute(attribute.Name) == null)
                    {
                        last.Add(content);
                    }
                    else
                    {
                        last.Attribute(attribute.Name).Value = attribute.Value;
                    }
                }
                else if (content != null)
                {
                    throw new NotSupportedException($"Unsupported content type {content.GetType().Name}: '{content}' for text formatting.");
                }
            }
        }

        internal Run GetFirstRunEffectedByEdit(int index, EditType type = EditType.Ins)
        {
            int len = HelperFunctions.GetText(Xml).Length;
            if (index < 0 || (type == EditType.Ins && index > len) || (type == EditType.Del && index >= len))
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            int count = 0;
            Run theOne = null;

            GetFirstRunEffectedByEditRecursive(Xml, index, ref count, ref theOne, type);

            return theOne;
        }

        internal void GetFirstRunEffectedByEditRecursive(XElement Xml, int index, ref int count, ref Run theOne, EditType type)
        {
            count += HelperFunctions.GetSize(Xml);

            // If the EditType is deletion then we must return the next blah
            if (count > 0 && ((type == EditType.Del && count > index) || (type == EditType.Ins && count >= index)))
            {
                // Correct the index
                foreach (XElement e in Xml.ElementsBeforeSelf())
                {
                    count -= HelperFunctions.GetSize(e);
                }

                count -= HelperFunctions.GetSize(Xml);

                // We have found the element, now find the run it belongs to.
                while ((Xml.Name.LocalName != "r") && (Xml.Name.LocalName != "pPr"))
                {
                    Xml = Xml.Parent;
                }

                theOne = new Run(Document, Xml, count);
            }
            else if (Xml.HasElements)
            {
                foreach (XElement e in Xml.Elements())
                {
                    if (theOne == null)
                    {
                        GetFirstRunEffectedByEditRecursive(e, index, ref count, ref theOne, type);
                    }
                }
            }
        }

        /// <summary>
        /// If the pPr element doesent exist it is created, either way it is returned by this function.
        /// </summary>
        /// <returns>The pPr element for this Paragraph.</returns>
        internal XElement ParaProperties()
        {
            // Get the element.
            XElement pPr = Xml.Element(DocxNamespace.Main + "pPr");

            // If it dosen't exist, create it.
            if (pPr == null)
            {
                Xml.AddFirst(new XElement(DocxNamespace.Main + "pPr"));
                pPr = Xml.Element(DocxNamespace.Main + "pPr");
            }

            // Return the pPr element for this Paragraph.
            return pPr;
        }

        /// <summary>
        /// If the ind element doesent exist it is created, either way it is returned by this function.
        /// </summary>
        /// <returns>The ind element for this Paragraphs pPr.</returns>
        internal XElement GetOrCreate_pPr_ind()
        {
            // Get the element.
            XElement pPr = ParaProperties();
            XElement ind = pPr.Element(DocxNamespace.Main + "ind");
            if (ind == null)
            {
                pPr.Add(new XElement(DocxNamespace.Main + "ind"));
                ind = pPr.Element(DocxNamespace.Main + "ind");
            }
            return ind;
        }

        internal string GetOrGenerateRel(Picture p)
        {
            string image_uri_string = p.img.pr.TargetUri.OriginalString;

            // Search for a relationship with a TargetUri that points at this Image.
            string Id = packagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                .Where(r => r.TargetUri.OriginalString == image_uri_string)
                .Select(r => r.Id)
                .SingleOrDefault();

            // If such a relation dosen't exist, create one.
            if (Id == null)
            {
                // Check to see if a relationship for this Picture exists and create it if not.
                PackageRelationship pr = packagePart.CreateRelationship(p.img.pr.TargetUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                Id = pr.Id;
            }
            return Id;
        }

        internal string GetOrGenerateRel(Hyperlink h)
        {
            string image_uri_string = h.Uri.OriginalString;

            // Search for a relationship with a TargetUri that points at this Image.
            string Id = packagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
                .Where(r => r.TargetUri.OriginalString == image_uri_string)
                .Select(r => r.Id)
                .SingleOrDefault();

            // If such a relation dosen't exist, create one.
            if (Id == null)
            {
                // Check to see if a relationship for this Picture exists and create it if not.
                PackageRelationship pr = packagePart.CreateRelationship(h.Uri, TargetMode.External, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
                Id = pr.Id;
            }
            return Id;
        }

        internal void RemoveHyperlinkRecursive(XElement xml, int index, ref int count, ref bool found)
        {
            if (xml.Name.LocalName.Equals("hyperlink", StringComparison.CurrentCultureIgnoreCase))
            {
                // This is the hyperlink to be removed.
                if (count == index)
                {
                    found = true;
                    xml.Remove();
                }
                else
                {
                    count++;
                }
            }

            if (xml.HasElements)
            {
                foreach (XElement e in xml.Elements())
                {
                    if (!found)
                    {
                        RemoveHyperlinkRecursive(e, index, ref count, ref found);
                    }
                }
            }
        }

        internal XElement[] SplitEdit(XElement edit, int index, EditType type)
        {
            Run run = GetFirstRunEffectedByEdit(index, type);
            XElement[] splitRun = Run.SplitRun(run, index, type);

            XElement splitLeft = new XElement(edit.Name, edit.Attributes(), run.Xml.ElementsBeforeSelf(), splitRun[0]);
            if (GetElementTextLength(splitLeft) == 0)
            {
                splitLeft = null;
            }

            XElement splitRight = new XElement(edit.Name, edit.Attributes(), splitRun[1], run.Xml.ElementsAfterSelf());
            if (GetElementTextLength(splitRight) == 0)
            {
                splitRight = null;
            }

            return new[] { splitLeft, splitRight };
        }
    }
}