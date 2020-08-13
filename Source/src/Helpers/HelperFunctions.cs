using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace DXPlus
{
    internal static class HelperFunctions
    {
        /// <summary>
        /// Checks whether 'toCheck' has all children that 'desired' has and values of 'val' attributes are the same
        /// </summary>
        /// <param name="desired"></param>
        /// <param name="toCheck"></param>
        /// <param name="formatOptions">Matching options whether check if desired attributes are inder a, or a has exactly and only these attributes as b has.</param>
        internal static bool ContainsEveryChildOf(XElement desired, XElement toCheck, MatchFormattingOptions formatOptions)
        {
            foreach (XElement e in desired.Elements())
            {
                // If a formatting property has the same name and 'val' attribute's value, its considered to be equivalent.
                if (!toCheck.Elements(e.Name).Any(bElement => bElement.GetVal() == e.GetVal()))
                {
                    return false;
                }
            }

            // If the formatting has to be exact, no additionaly formatting must exist.
            return formatOptions != MatchFormattingOptions.ExactMatch
                || desired.Elements().Count() == toCheck.Elements().Count();
        }

        internal static void CreateRelsPackagePart(DocX Document, Uri uri)
        {
            PackagePart pp = Document.package.CreatePart(uri, DocxContentType.Relationships, CompressionOption.Maximum);
            using TextWriter tw = new StreamWriter(pp.GetStream());
            XDocument d = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(DocxNamespace.RelatedPackage + "Relationships")
            );

            d.Save(tw);
        }

        internal static int GetSize(XElement Xml)
        {
            switch (Xml.Name.LocalName)
            {
                case "tab":
                case "br":
                case "tr":
                case "tc":
                    return 1;

                case "t":
                case "delText":
                    return Xml.Value.Length;

                default:
                    return 0;
            }
        }

        internal static string GetText(XElement e) => GetTextRecursive(e).ToString();

        internal static StringBuilder GetTextRecursive(XElement Xml, StringBuilder sb = null)
        {
            (sb ??= new StringBuilder()).Append(ToText(Xml));

            if (Xml.HasElements)
            {
                foreach (XElement e in Xml.Elements())
                {
                    GetTextRecursive(e, sb);
                }
            }

            return sb;
        }

        internal static List<FormattedText> GetFormattedText(XElement e)
        {
            List<FormattedText> alist = new List<FormattedText>();
            GetFormattedTextRecursive(e, ref alist);
            return alist;
        }

        internal static void GetFormattedTextRecursive(XElement Xml, ref List<FormattedText> alist)
        {
            FormattedText ft = ToFormattedText(Xml);
            FormattedText last = null;

            if (ft != null)
            {
                if (alist.Count > 0)
                    last = alist.Last();

                if (last?.CompareTo(ft) == 0)
                {
                    last.text += ft.text;
                }
                else
                {
                    if (last != null)
                    {
                        ft.index = last.index + last.text.Length;
                    }

                    alist.Add(ft);
                }
            }

            if (Xml.HasElements)
            {
                foreach (XElement e in Xml.Elements())
                {
                    GetFormattedTextRecursive(e, ref alist);
                }
            }
        }

        internal static FormattedText ToFormattedText(XElement e)
        {
            // The text representation of e.
            string text = ToText(e);
            if (string.IsNullOrEmpty(text))
                return null;

            // e is a w:t element, it must exist inside a w:r element or a w:tabs, lets climb until we find it.
            while (!e.Name.Equals(DocxNamespace.Main + "r") &&
                   !e.Name.Equals(DocxNamespace.Main + "tabs"))
            {
                e = e.Parent;
            }

            // e is a w:r element, lets find the rPr element.
            XElement rPr = e.Element(DocxNamespace.Main + "rPr");

            FormattedText ft = new FormattedText
            {
                text = text,
                index = 0,
                formatting = null
            };

            // Return text with formatting.
            if (rPr != null)
                ft.formatting = Formatting.Parse(rPr);

            return ft;
        }

        internal static string ToText(XElement e)
        {
            switch (e.Name.LocalName)
            {
                case "tab":
                case "tc":
                    return "\t";

                case "tr":
                case "br":
                    return "\n";

                case "t":
                case "delText":
                    {
                        if (e.Parent?.Name.LocalName == "r")
                        {
                            XElement run = e.Parent;
                            var rPr = run.Elements().FirstOrDefault(a => a.Name.LocalName == "rPr");
                            if (rPr != null)
                            {
                                var caps = rPr.Elements().FirstOrDefault(a => a.Name.LocalName == "caps");
                                if (caps != null)
                                    return e.Value.ToUpper();
                            }
                        }

                        return e.Value;
                    }
                default:
                    return string.Empty;
            }
        }

        internal static XElement CloneElement(XElement element)
        {
            return new XElement(
                element.Name,
                element.Attributes(),
                element.Nodes().Select(n => n is XElement e ? CloneElement(e) : n)
            );
        }

        internal static PackagePart CreateOrGetSettingsPart(Package package)
        {
            PackagePart settingsPart;

            if (!package.PartExists(DocxSections.SettingsUri))
            {
                settingsPart = package.CreatePart(DocxSections.SettingsUri, DocxContentType.Settings, CompressionOption.Maximum);

                PackagePart mainDocumentPart = package.GetParts().Single(p =>
                       p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
                    || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

                mainDocumentPart.CreateRelationship(DocxSections.SettingsUri, TargetMode.Internal, $"{DocxNamespace.RelatedDoc.NamespaceName}/settings");

                XDocument settings = XDocument.Parse
                (@"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                <w:settings xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w10='urn:schemas-microsoft-com:office:word' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:sl='http://schemas.openxmlformats.org/schemaLibrary/2006/main'>
                  <w:zoom w:percent='100' />
                  <w:defaultTabStop w:val='720' />
                  <w:characterSpacingControl w:val='doNotCompress' />
                  <w:compat />
                  <w:rsids>
                    <w:rsidRoot w:val='00217F62' />
                    <w:rsid w:val='001915A3' />
                    <w:rsid w:val='00217F62' />
                    <w:rsid w:val='00A906D8' />
                    <w:rsid w:val='00AB5A74' />
                    <w:rsid w:val='00F071AE' />
                  </w:rsids>
                  <m:mathPr>
                    <m:mathFont m:val='Cambria Math' />
                    <m:brkBin m:val='before' />
                    <m:brkBinSub m:val='--' />
                    <m:smallFrac m:val='off' />
                    <m:dispDef />
                    <m:lMargin m:val='0' />
                    <m:rMargin m:val='0' />
                    <m:defJc m:val='centerGroup' />
                    <m:wrapIndent m:val='1440' />
                    <m:intLim m:val='subSup' />
                    <m:naryLim m:val='undOvr' />
                  </m:mathPr>
                  <w:themeFontLang w:val='en-IE' w:bidi='ar-SA' />
                  <w:clrSchemeMapping w:bg1='light1' w:t1='dark1' w:bg2='light2' w:t2='dark2' w:accent1='accent1' w:accent2='accent2' w:accent3='accent3' w:accent4='accent4' w:accent5='accent5' w:accent6='accent6' w:hyperlink='hyperlink' w:followedHyperlink='followedHyperlink' />
                  <w:shapeDefaults>
                    <o:shapedefaults v:ext='edit' spidmax='2050' />
                    <o:shapelayout v:ext='edit'>
                      <o:idmap v:ext='edit' data='1' />
                    </o:shapelayout>
                  </w:shapeDefaults>
                  <w:decimalSymbol w:val='.' />
                  <w:listSeparator w:val=',' />
                </w:settings>"
                );

                XElement themeFontLang = settings.Root.Element(DocxNamespace.Main + "themeFontLang");
                themeFontLang.SetVal(CultureInfo.CurrentCulture);

                // Save the settings document.
                using TextWriter tw = new StreamWriter(settingsPart.GetStream());
                settings.Save(tw);
            }
            else
            {
                settingsPart = package.GetPart(DocxSections.SettingsUri);
            }

            return settingsPart;
        }

        internal static void CreateCustomPropertiesPart(DocX document)
        {
            PackagePart customPropertiesPart = document.package.CreatePart(new Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml", CompressionOption.Maximum);

            XDocument customPropDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(DocxNamespace.CustomPropertiesSchema + "Properties",
                    new XAttribute(XNamespace.Xmlns + "vt", DocxNamespace.CustomVTypesSchema)
                )
            );

            using (TextWriter tw = new StreamWriter(customPropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
                customPropDoc.Save(tw, SaveOptions.None);

            document.package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal,
                $"{DocxNamespace.RelatedDoc.NamespaceName}/custom-properties");
        }

        /// <summary>
        /// If this document does not contain a /word/numbering.xml add the default one generated by Microsoft Word
        /// when the default bullet, numbered and multilevel lists are added to a blank document
        /// </summary>
        /// <param name="package"></param>
        internal static XDocument AddDefaultNumberingXml(Package package)
        {
            PackagePart wordNumbering = package.CreatePart(new Uri("/word/numbering.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", CompressionOption.Maximum);
            var numberingDoc = Resources.NumberingXml;

            // Save /word/numbering.xml
            using (TextWriter tw = new StreamWriter(wordNumbering.GetStream(FileMode.Create, FileAccess.Write)))
                numberingDoc.Save(tw, SaveOptions.None);

            PackagePart mainDocumentPart = package.GetParts().Single(p =>
                p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
                || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

            mainDocumentPart.CreateRelationship(wordNumbering.Uri, TargetMode.Internal, $"{DocxNamespace.RelatedDoc.NamespaceName}/numbering");
            return numberingDoc;
        }

        /// <summary>
        /// If this document does not contain a /word/styles.xml add the default one generated by Microsoft Word.
        /// </summary>
        /// <param name="package"></param>
        internal static XDocument AddDefaultStylesXml(Package package)
        {
            PackagePart word_styles = package.CreatePart(DocxSections.StylesUri, DocxContentType.Styles, CompressionOption.Maximum);
            var stylesDoc = Resources.DefaultStylesXml;

            XElement lang = stylesDoc.Root.Element(DocxNamespace.Main + "docDefaults")
                                     .Element(DocxNamespace.Main + "rPrDefault")
                                     .Element(DocxNamespace.Main + "rPr")
                                     .Element(DocxNamespace.Main + "lang");
            lang.SetAttributeValue(DocxNamespace.Main + "val", CultureInfo.CurrentCulture);

            // Save /word/styles.xml
            using (TextWriter tw = new StreamWriter(word_styles.GetStream(FileMode.Create, FileAccess.Write)))
                stylesDoc.Save(tw, SaveOptions.None);

            PackagePart mainDocumentPart = package.GetParts().Single(p =>
                p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
                || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase)
            );

            mainDocumentPart.CreateRelationship(word_styles.Uri, TargetMode.Internal, $"{DocxNamespace.RelatedDoc.NamespaceName}/styles");
            return stylesDoc;
        }

        /// <summary>
        /// Creates an Edit either a ins or a del with the specified content and date
        /// </summary>
        /// <param name="t">The type of this edit (ins or del)</param>
        /// <param name="edit_time">The time stamp to use for this edit</param>
        /// <param name="content">The initial content of this edit</param>
        internal static XElement CreateEdit(EditType editType, DateTime editTime, object content)
        {
            if (editType == EditType.Del && content is IEnumerable<XElement> iex)
            {
                foreach (var e in iex)
                {
                    IEnumerable<XElement> ts = e.DescendantsAndSelf(DocxNamespace.Main + "t");

                    for (int i = 0; i < ts.Count(); i++)
                    {
                        XElement text = ts.ElementAt(i);
                        text.ReplaceWith(new XElement(DocxNamespace.Main + "delText", text.Attributes(), text.Value));
                    }
                }
            }

            return new XElement(DocxNamespace.Main + editType.ToString(),
                    new XAttribute(DocxNamespace.Main + "id", 0),
                    new XAttribute(DocxNamespace.Main + "author", Environment.UserName),
                    new XAttribute(DocxNamespace.Main + "date", editTime),
                    content);
        }

        internal static XElement CreateTable(int rowCount, int columnCount)
        {
            int[] columnWidths = new int[columnCount];
            for (int i = 0; i < columnCount; i++)
            {
                columnWidths[i] = 2310;
            }
            return CreateTable(rowCount, columnWidths);
        }

        internal static XElement CreateTable(int rowCount, int[] columnWidths)
        {
            XElement newTable = new XElement(
                DocxNamespace.Main + "tbl",
                new XElement
                (
                    DocxNamespace.Main + "tblPr",
                        new XElement(DocxNamespace.Main + "tblStyle", new XAttribute(DocxNamespace.Main + "val", "TableGrid")),
                        new XElement(DocxNamespace.Main + "tblW", new XAttribute(DocxNamespace.Main + "w", "5000"), new XAttribute(DocxNamespace.Main + "type", "auto")),
                        new XElement(DocxNamespace.Main + "tblLook", new XAttribute(DocxNamespace.Main + "val", "04A0"))
                )
            );

            for (int i = 0; i < rowCount; i++)
            {
                XElement row = new XElement(DocxNamespace.Main + "tr");

                for (int j = 0; j < columnWidths.Length; j++)
                {
                    XElement cell = CreateTableCell();
                    row.Add(cell);
                }

                newTable.Add(row);
            }
            return newTable;
        }

        /// <summary>
        /// Create and return a cell of a table
        /// </summary>
        internal static XElement CreateTableCell(double w = 2310) => new XElement(
                DocxNamespace.Main + "tc",
                    new XElement(DocxNamespace.Main + "tcPr",
                    new XElement(DocxNamespace.Main + "tcW",
                            new XAttribute(DocxNamespace.Main + "w", w),
                            new XAttribute(DocxNamespace.Main + "type", "dxa"))),
                    new XElement(DocxNamespace.Main + "p",
                        new XElement(DocxNamespace.Main + "pPr"))
            );

        internal static List CreateItemInList(List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false)
        {
            if (list.NumId == 0)
            {
                list.CreateNewNumberingNumId(listType, startNumber, continueNumbering);
            }

            if (listText != null)
            {
                var newParagraphSection = new XElement(
                    DocxNamespace.Main + "p",
                    new XElement(DocxNamespace.Main + "pPr",
                                 new XElement(DocxNamespace.Main + "numPr",
                                              new XElement(DocxNamespace.Main + "ilvl", new XAttribute(DocxNamespace.Main + "val", level)),
                                              new XElement(DocxNamespace.Main + "numId", new XAttribute(DocxNamespace.Main + "val", list.NumId)))),
                    new XElement(DocxNamespace.Main + "r", new XElement(DocxNamespace.Main + "t", listText))
                );

                if (trackChanges)
                    newParagraphSection = CreateEdit(EditType.Ins, DateTime.Now, newParagraphSection);

                if (startNumber == null)
                {
                    list.AddItem(new Paragraph(list.Document, newParagraphSection, 0, ContainerType.Paragraph));
                }
                else
                {
                    list.AddItem(new Paragraph(list.Document, newParagraphSection, 0, ContainerType.Paragraph), (int)startNumber);
                }
            }

            return list;
        }

        internal static void RenumberIDs(DocX document)
        {
            var trackerIDs = document.mainDoc.Descendants()
                    .Where(d => d.Name.LocalName == "ins" || d.Name.LocalName == "del")
                    .Select(d => d.Attribute(DocxNamespace.Main + "id"))
                    .ToList();

            for (int i = 0; i < trackerIDs.Count; i++)
                trackerIDs[i].Value = i.ToString();
        }

        internal static Paragraph GetFirstParagraphAffectedByInsert(DocX document, int index)
        {
            // This document contains no Paragraphs and insertion is at index 0
            if (document.paragraphLookup.Keys.Count == 0 && index == 0)
                return null;

            foreach (int paragraphEndIndex in document.paragraphLookup.Keys)
            {
                if (paragraphEndIndex >= index)
                    return document.paragraphLookup[paragraphEndIndex];
            }

            throw new ArgumentOutOfRangeException();
        }

        internal static List<XElement> FormatInput(string text, XElement rPr)
        {
            List<XElement> newRuns = new List<XElement>();
            XElement tabRun = new XElement(DocxNamespace.Main + "tab");
            XElement breakRun = new XElement(DocxNamespace.Main + "br");

            StringBuilder sb = new StringBuilder();

            if (string.IsNullOrEmpty(text))
            {
                return newRuns;
            }

            foreach (char c in text)
            {
                switch (c)
                {
                    case '\t':
                        if (sb.Length > 0)
                        {
                            XElement t = new XElement(DocxNamespace.Main + "t", sb.ToString());
                            TextBlock.PreserveSpace(t);
                            newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr, t));
                            sb = new StringBuilder();
                        }
                        newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr, tabRun));
                        break;

                    case '\r':
                    case '\n':
                        if (sb.Length > 0)
                        {
                            XElement t = new XElement(DocxNamespace.Main + "t", sb.ToString());
                            TextBlock.PreserveSpace(t);
                            newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr, t));
                            sb = new StringBuilder();
                        }
                        newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr, breakRun));
                        break;

                    default:
                        sb.Append(c);
                        break;
                }
            }

            if (sb.Length > 0)
            {
                XElement t = new XElement(DocxNamespace.Main + "t", sb.ToString());
                TextBlock.PreserveSpace(t);
                newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr, t));
            }

            return newRuns;
        }

        internal static XElement[] SplitParagraph(Paragraph p, int index)
        {
            // In this case edit dosent really matter, you have a choice.
            Run r = p.GetFirstRunEffectedByEdit(index, EditType.Ins);

            XElement[] split;
            XElement before, after;

            if (r.Xml.Parent.Name.LocalName == "ins")
            {
                split = p.SplitEdit(r.Xml.Parent, index, EditType.Ins);
                before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
            }
            else if (r.Xml.Parent.Name.LocalName == "del")
            {
                split = p.SplitEdit(r.Xml.Parent, index, EditType.Del);

                before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
            }
            else
            {
                split = Run.SplitRun(r, index);

                before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.ElementsBeforeSelf(), split[0]);
                after = new XElement(p.Xml.Name, p.Xml.Attributes(), split[1], r.Xml.ElementsAfterSelf());
            }

            if (!before.Elements().Any())
                before = null;

            if (!after.Elements().Any())
                after = null;

            return new XElement[] { before, after };
        }

        internal static bool IsSameFile(Stream streamOne, Stream streamTwo)
        {
            int file1byte, file2byte;

            if (streamOne.Length != streamTwo.Length)
            {
                return false;
            }

            // Read and compare a byte from each file until either a
            // non-matching set of bytes is found or until the end of
            // file1 is reached.
            do
            {
                // Read one byte from each file.
                file1byte = streamOne.ReadByte();
                file2byte = streamTwo.ReadByte();
            }
            while ((file1byte == file2byte) && (file1byte != -1));

            // Return the success of the comparison. "file1byte" is
            // equal to "file2byte" at this point only if the files are
            // the same.

            streamOne.Position = 0;
            streamTwo.Position = 0;

            return (file1byte - file2byte) == 0;
        }

        internal static byte[] ConcatByteArrays(byte[] array1, byte[] array2)
        {
            byte[] result = new byte[array1.Length + array2.Length];
            Buffer.BlockCopy(array2, 0, result, 0, array2.Length);
            Buffer.BlockCopy(array1, 0, result, array2.Length, array1.Length);
            return result;
        }
    }
}