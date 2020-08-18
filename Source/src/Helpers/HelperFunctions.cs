using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    internal static class HelperFunctions
    {
        /// <summary>
        /// Method to persist XML back to a PackagePart
        /// </summary>
        /// <param name="packagePart"></param>
        /// <param name="document"></param>
        internal static void Save(this PackagePart packagePart, XDocument document)
        {
            if (packagePart == null)
                throw new ArgumentNullException(nameof(packagePart));
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            lock (packagePart.Package)
            {
                using var writer = new StreamWriter(packagePart.GetStream(FileMode.OpenOrCreate, FileAccess.Write), Encoding.UTF8);
                document.Save(writer, SaveOptions.OmitDuplicateNamespaces);
            }

#if DEBUG
            _ = Load(packagePart);
#endif
        }

        /// <summary>
        /// Load XML from a package part
        /// </summary>
        /// <param name="packagePart"></param>
        /// <returns></returns>
        internal static XDocument Load(this PackagePart packagePart)
        {
            if (packagePart == null)
                throw new ArgumentNullException(nameof(packagePart));

            lock (packagePart.Package)
            {
                using var reader = new StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8);
                var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
                if (document.Root == null)
                    throw new Exception("Loaded document from PackagePart has no contents.");

                return document;
            }
        }

        /// <summary>
        /// Checks whether 'toCheck' has all children that 'desired' has and values of 'val' attributes are the same
        /// </summary>
        internal static bool ContainsEveryChildOf(XElement desired, XElement toCheck, MatchFormattingOptions formatOptions)
        {
            if (desired == null)
                throw new ArgumentNullException(nameof(desired));
            if (toCheck == null) 
                throw new ArgumentNullException(nameof(toCheck));

            if (desired.Elements().Any(e => toCheck.Elements(e.Name)
                .All(bElement => bElement.GetVal() != e.GetVal())))
            {
                return false;
            }

            // If the formatting has to be exact, no additional formatting must exist.
            return formatOptions != MatchFormattingOptions.ExactMatch
                || desired.Elements().Count() == toCheck.Elements().Count();
        }

        internal static void CreateRelsPackagePart(DocX document, Uri uri)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (uri == null)
                throw new ArgumentNullException(nameof(uri));

            PackagePart pp = document.Package.CreatePart(uri, DocxContentType.Relationships, CompressionOption.Maximum);
            using TextWriter tw = new StreamWriter(pp.GetStream());
            XDocument d = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(DocxNamespace.RelatedPackage + "Relationships")
            );

            d.Save(tw);
        }

        internal static int GetSize(XElement xml)
        {
            if (xml == null) 
                throw new ArgumentNullException(nameof(xml));

            switch (xml.Name.LocalName)
            {
                case "tab":
                case "br":
                case "tr":
                case "tc":
                    return 1;

                case "t":
                case "delText":
                    return xml.Value.Length;

                default:
                    return 0;
            }
        }

        internal static string GetText(XElement e)
        {
            return GetTextRecursive(e).ToString();
        }

        internal static StringBuilder GetTextRecursive(XElement xml, StringBuilder sb = null)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));
            
            (sb ??= new StringBuilder()).Append(ToText(xml));

            if (xml.HasElements)
            {
                foreach (var e in xml.Elements())
                {
                    GetTextRecursive(e, sb);
                }
            }

            return sb;
        }

        internal static List<FormattedText> GetFormattedText(XElement xml)
        {
           
            var list = new List<FormattedText>();
            GetFormattedTextRecursive(xml, ref list);
            return list;
        }

        internal static void GetFormattedTextRecursive(XElement xml, ref List<FormattedText> list)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));

            var ft = ToFormattedText(xml);
            FormattedText last = null;

            if (ft != null)
            {
                if (list.Count > 0)
                {
                    last = list.Last();
                }

                if (last?.CompareTo(ft) == 0)
                {
                    last.Text += ft.Text;
                }
                else
                {
                    if (last != null)
                    {
                        ft.Index = last.Index + last.Text.Length;
                    }

                    list.Add(ft);
                }
            }

            if (xml.HasElements)
            {
                foreach (var e in xml.Elements())
                {
                    GetFormattedTextRecursive(e, ref list);
                }
            }
        }

        internal static FormattedText ToFormattedText(XElement xml)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));
            
            string text = ToText(xml);
            if (string.IsNullOrEmpty(text))
                return null;

            // xml is a w:t element, it must exist inside a w:r element or a w:tabs, lets climb until we find it.
            while (!xml.Name.Equals(DocxNamespace.Main + "r") &&
                   !xml.Name.Equals(DocxNamespace.Main + "tabs"))
            {
                xml = xml.Parent;
            }

            // xml is a w:r element, use the run properties.
            return new FormattedText {
                Text = text,
                Index = 0,
                Formatting = Formatting.Parse(xml.GetRunProps(false))
            };
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
                            var caps = e.Parent.FirstLocalNameDescendant("rPr")?
                                                       .FirstLocalNameDescendant("caps");
                            if (caps != null)
                                return e.Value.ToUpper();
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

                var mainDocumentPart = package.GetParts().Single(p =>
                       p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
                    || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

                mainDocumentPart.CreateRelationship(DocxSections.SettingsUri, TargetMode.Internal, $"{DocxNamespace.RelatedDoc.NamespaceName}/settings");

                var settings = XDocument.Parse
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

                settings.Root?.Element(DocxNamespace.Main + "themeFontLang")?.SetAttributeValue(DocxNamespace.Main + "val", CultureInfo.CurrentCulture);

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

        /// <summary>
        /// If this document does not contain a /word/styles.xml add the default one generated by Microsoft Word.
        /// </summary>
        /// <param name="package"></param>
        internal static XDocument AddDefaultStylesXml(Package package)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));
            if (package.PartExists(DocxSections.StylesUri))
                throw new InvalidOperationException("Root style collection already exists.");

            var wordStyles = package.CreatePart(DocxSections.StylesUri, DocxContentType.Styles, CompressionOption.Maximum);
            var stylesDoc = Resources.DefaultStylesXml;

            // Set the run default language to be the current culture.
            stylesDoc.Root.Element(DocxNamespace.Main + "docDefaults")
                          .Element(DocxNamespace.Main + "rPrDefault")
                          .Element(DocxNamespace.Main + "rPr")
                          .Element(DocxNamespace.Main + "lang")
                          .SetAttributeValue(DocxNamespace.Main + "val", CultureInfo.CurrentCulture);

            // Save /word/styles.xml
            wordStyles.Save(stylesDoc);

            var mainDocumentPart = package.GetParts().Single(p =>
                p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
                || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase)
            );

            mainDocumentPart.CreateRelationship(wordStyles.Uri, TargetMode.Internal, $"{DocxNamespace.RelatedDoc.NamespaceName}/styles");
            return stylesDoc;
        }

        /// <summary>
        /// Creates an Edit either a ins or a del with the specified content and date
        /// </summary>
        /// <param name="editType">The type of this edit (ins or del)</param>
        /// <param name="editTime">The time stamp to use for this edit</param>
        /// <param name="content">The initial content of this edit</param>
        internal static XElement CreateEdit(EditType editType, DateTime editTime, object content)
        {
            if (editType == EditType.Del && content is IEnumerable<XElement> iex)
            {
                foreach (var e in iex)
                {
                    var ts = e.DescendantsAndSelf(DocxNamespace.Main + "t").ToList();
                    foreach (var text in ts)
                    {
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

        private const int DEFAULT_COL_WIDTH = 2310;

        internal static XElement CreateTable(int rowCount, int columnCount)
        {
            int[] columnWidths = new int[columnCount];
            Array.Fill(columnWidths, DEFAULT_COL_WIDTH);

            return CreateTable(rowCount, columnWidths);
        }

        internal static XElement CreateTable(int rowCount, int[] columnWidths)
        {
            var newTable = new XElement(
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
                var row = new XElement(DocxNamespace.Main + "tr");
                for (int j = 0; j < columnWidths.Length; j++)
                {
                    var cell = CreateTableCell();
                    row.Add(cell);
                }

                newTable.Add(row);
            }
            return newTable;
        }

        /// <summary>
        /// Create and return a cell of a table
        /// </summary>
        internal static XElement CreateTableCell(double width = DEFAULT_COL_WIDTH)
        {
            return new XElement(DocxNamespace.Main + "tc",
                    new XElement(DocxNamespace.Main + "tcPr",
                        new XElement(DocxNamespace.Main + "tcW",
                            new XAttribute(DocxNamespace.Main + "w", width),
                            new XAttribute(DocxNamespace.Main + "type", "dxa"))),
                    new XElement(DocxNamespace.Main + "p",
                    new XElement(DocxNamespace.Main + "pPr"))
            );
        }

        internal static List CreateItemInList(List list, string listText, int level = 0, ListItemType listType = ListItemType.Numbered, int? startNumber = null, bool trackChanges = false, bool continueNumbering = false)
        {
            if (list == null)
                throw new ArgumentNullException(nameof(list));
            
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
                {
                    newParagraphSection = CreateEdit(EditType.Ins, DateTime.Now, newParagraphSection);
                }

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

        internal static void RenumberIds(DocX document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            var trackerIds = document.mainDoc.Descendants()
                    .Where(d => d.Name.LocalName == "ins" || d.Name.LocalName == "del")
                    .Select(d => d.Attribute(DocxNamespace.Main + "id"))
                    .ToList();

            for (int i = 0; i < trackerIds.Count; i++)
            {
                trackerIds[i].Value = i.ToString();
            }
        }

        internal static Paragraph GetFirstParagraphAffectedByInsert(DocX document, int index)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            // If the insertion position is first (0) and there are no paragraphs, then return null.
            if (document.paragraphLookup.Keys.Count == 0 && index == 0)
                return null;

            // Find the paragraph that contains the index
            foreach (var paragraphEndIndex in document.paragraphLookup.Keys.Where(paragraphEndIndex => paragraphEndIndex >= index))
            {
                return document.paragraphLookup[paragraphEndIndex];
            }

            throw new ArgumentOutOfRangeException(nameof(index));
        }

        internal static List<XElement> FormatInput(string text, XElement rPr)
        {
            var newRuns = new List<XElement>();
            if (string.IsNullOrEmpty(text))
            {
                return newRuns;
            }

            var tabRun = new XElement(DocxNamespace.Main + "tab");
            var breakRun = new XElement(DocxNamespace.Main + "br");
            var sb = new StringBuilder();

            foreach (var c in text)
            {
                switch (c)
                {
                    case '\t':
                        if (sb.Length > 0)
                        {
                            newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr,
                                new XElement(DocxNamespace.Main + "t", sb.ToString()).PreserveSpace()));
                            sb = new StringBuilder();
                        }
                        newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr, tabRun));
                        break;

                    case '\r':
                    case '\n':
                        if (sb.Length > 0)
                        {
                            newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr,
                                new XElement(DocxNamespace.Main + "t", sb.ToString()).PreserveSpace()));
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
                newRuns.Add(new XElement(DocxNamespace.Main + "r", rPr,
                    new XElement(DocxNamespace.Main + "t", sb.ToString()).PreserveSpace()));
            }

            return newRuns;
        }

        internal static XElement[] SplitParagraph(Paragraph p, int index)
        {
            if (p == null)
                throw new ArgumentNullException(nameof(p));
            
            var r = p.GetFirstRunEffectedByEdit(index);
            XElement[] split;
            XElement before, after;

            switch (r.Xml.Parent?.Name.LocalName)
            {
                case "ins":
                    split = p.SplitEdit(r.Xml.Parent, index, EditType.Ins);
                    before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                    after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
                    break;
                case "del":
                    split = p.SplitEdit(r.Xml.Parent, index, EditType.Del);
                    before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                    after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
                    break;
                default:
                    split = Run.SplitRun(r, index);
                    before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.ElementsBeforeSelf(), split[0]);
                    after = new XElement(p.Xml.Name, p.Xml.Attributes(), split[1], r.Xml.ElementsAfterSelf());
                    break;
            }

            if (!before.Elements().Any())
                before = null;

            if (!after.Elements().Any())
                after = null;

            return new[] { before, after };
        }

        internal static bool IsSameFile(Stream streamOne, Stream streamTwo)
        {
            if (streamOne == null)
                throw new ArgumentNullException(nameof(streamOne));
            if (streamTwo == null)
                throw new ArgumentNullException(nameof(streamTwo));
            
            if (streamOne.Length != streamTwo.Length)
                return false;

            int b1, b2;
            do
            {
                // Read one byte from each file.
                b1 = streamOne.ReadByte();
                b2 = streamTwo.ReadByte();
            }
            while (b1 == b2 && b1 != -1);

            streamOne.Position = 0;
            streamTwo.Position = 0;
           
            return b1 == b2;
        }

        internal static byte[] ConcatByteArrays(byte[] array1, byte[] array2)
        {
            if (array1 == null)
                throw new ArgumentNullException(nameof(array1));
            if (array2 == null)
                throw new ArgumentNullException(nameof(array2));
            
            byte[] result = new byte[array1.Length + array2.Length];
            Buffer.BlockCopy(array2, 0, result, 0, array2.Length);
            Buffer.BlockCopy(array1, 0, result, array2.Length, array1.Length);
            
            return result;
        }
    }
}