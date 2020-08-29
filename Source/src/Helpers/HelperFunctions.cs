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
        /// This creates a Word docx in a memory stream.
        /// </summary>
        /// <param name="documentType">Type (doc or template)</param>
        /// <returns>Memory stream with loaded doc</returns>
        internal static Stream CreateDocumentType(DocumentTypes documentType)
        {
            // Create the docx package
            var ms = new MemoryStream();
            using var package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);

            // Create the main document part for this package
            var mainDocumentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative),
                documentType == DocumentTypes.Document ? DocxContentType.Document : DocxContentType.Template,
                CompressionOption.Normal);
            package.CreateRelationship(mainDocumentPart.Uri, TargetMode.Internal, $"{Namespace.RelatedDoc.NamespaceName}/officeDocument");

            // Generate an id for this editing session.
            var startingRevisionId = GenerateRevisionStamp(null, out _);

            // Load the document part into a XDocument object
            var mainDoc = Resources.BodyDocument(startingRevisionId);

            // Add the settings.xml + relationship
            AddDefaultSettingsPart(package, startingRevisionId);

            // Add the default styles + relationship
            AddDefaultStylesXml(package, out var stylesPart, out var stylesDoc);

            // Save the main new document back to the package.
            mainDocumentPart.Save(mainDoc);
            package.Close();

            // Return the stream representing the created document.
            return ms;
        }

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

        internal static bool IsValidHexNumber(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                    && value.Length <= 8
                    && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int _);
        }

        internal static int GetSize(XElement xml)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));

            switch (xml.Name.LocalName)
            {
                case "tab":
                case "cr":
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
            return GetTextRecursive(e)?.ToString() ?? string.Empty;
        }

        internal static StringBuilder GetTextRecursive(XElement xml, StringBuilder sb = null)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));

            string text = ToText(xml);
            if (!string.IsNullOrEmpty(text))
            {
                (sb ??= new StringBuilder()).Append(text);
            }

            if (xml.HasElements)
            {
                sb = xml.Elements().Aggregate(sb, (current, e) => GetTextRecursive(e, current));
            }

            return sb;
        }

        /*
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
            while (!xml.Name.Equals(Name.Run) &&
                   !xml.Name.Equals(Namespace.Main + "tabs"))
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
        */

        /// <summary>
        /// Turn a Word (w:t) element into text.
        /// </summary>
        /// <param name="e"></param>
        internal static string ToText(XElement e)
        {
            switch (e.Name.LocalName)
            {
                case "tab":
                case "tc":
                    return "\t";

                case "cr":
                case "tr":
                case "br":
                    return "\n";

                case "t":
                case "delText":
                    if (e.Parent?.Name == Name.Run)
                    {
                        // Get the caps setting.
                        var props = new Formatting(e.Parent!.Element(Name.RunProperties));
                        if (props.CapsStyle != CapsStyle.None)
                            return e.Value.ToUpper();
                    }
                    return e.Value;
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Clone an XElement into a new object
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        internal static XElement Clone(this XElement element)
        {
            return new XElement(
                element.Name,
                element.Attributes(),
                element.Nodes().Select(n => n is XElement e ? Clone(e) : n)
            );
        }

        /// <summary>
        /// If this document does not contain a /word/styles.xml add the default one generated by Microsoft Word.
        /// </summary>
        /// <param name="package"></param>
        /// <param name="stylesPart"></param>
        /// <param name="stylesDoc"></param>
        internal static void AddDefaultStylesXml(Package package, out PackagePart stylesPart, out XDocument stylesDoc)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));
            if (package.PartExists(Relations.Styles.Uri))
                throw new InvalidOperationException("Root style collection already exists.");

            stylesPart = package.CreatePart(Relations.Styles.Uri, Relations.Styles.ContentType, CompressionOption.Maximum);
            stylesDoc = Resources.DefaultStylesXml();

            // Set the run default language to be the current culture.
            stylesDoc.Root!.Element(Namespace.Main + "docDefaults")
                .Element(Namespace.Main + "rPrDefault")
                .Element(Name.RunProperties)
                .Element(Name.Language)
                .SetAttributeValue(Name.MainVal, CultureInfo.CurrentCulture);

            // Save /word/styles.xml
            stylesPart.Save(stylesDoc);

            // Add the relationship to the main doc
            var mainDocumentPart = package.GetParts().Single(p =>
                p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
                || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

            mainDocumentPart.CreateRelationship(stylesPart.Uri, TargetMode.Internal, $"{Namespace.RelatedDoc.NamespaceName}/styles");
        }

        /// <summary>
        /// Create the /word/settings.xml document
        /// </summary>
        /// <param name="package">Package owner</param>
        /// <param name="rsid">Initial document revision id</param>
        internal static void AddDefaultSettingsPart(Package package, string rsid)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));
            if (string.IsNullOrEmpty(rsid))
                throw new ArgumentException($"'{nameof(rsid)}' cannot be null or empty", nameof(rsid));
            if (package.PartExists(Relations.Settings.Uri))
                throw new InvalidOperationException("Settings.xml section already exists.");

            // Add the settings package part and document
            var settingsPart = package.CreatePart(Relations.Settings.Uri, Relations.Settings.ContentType, CompressionOption.Maximum);
            var settings = Resources.SettingsXml(rsid);

            // Set the correct language
            settings.Root.Element(Namespace.Main + "themeFontLang")
                         .SetAttributeValue(Name.MainVal, CultureInfo.CurrentCulture);

            // Save the settings document.
            settingsPart.Save(settings);

            // Add the relationship to the main doc
            var mainDocumentPart = package.GetParts().Single(p =>
                    p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
                 || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

            mainDocumentPart.CreateRelationship(Relations.Settings.Uri, TargetMode.Internal, Relations.Settings.RelType);
        }

        internal static Paragraph GetFirstParagraphAffectedByInsert(DocX document, int index)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            // If the insertion position is first (0) and there are no paragraphs, then return null.
            var lookup = document.GetParagraphIndexes();
            if (lookup.Keys.Count == 0 && index == 0)
                return null;

            // Find the paragraph that contains the index
            foreach (var paragraphEndIndex in lookup.Keys.Where(paragraphEndIndex => paragraphEndIndex >= index))
            {
                return lookup[paragraphEndIndex];
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

            var tabRun = new XElement(Namespace.Main + "tab");
            var breakRun = new XElement(Namespace.Main + "br");
            var sb = new StringBuilder();

            foreach (var c in text)
            {
                switch (c)
                {
                    case '\t':
                        if (sb.Length > 0)
                        {
                            newRuns.Add(new XElement(Name.Run, rPr,
                                new XElement(Name.Text, sb.ToString()).PreserveSpace()));
                            sb = new StringBuilder();
                        }
                        newRuns.Add(new XElement(Name.Run, rPr, tabRun));
                        break;

                    case '\r':
                    case '\n':
                        if (sb.Length > 0)
                        {
                            newRuns.Add(new XElement(Name.Run, rPr,
                                new XElement(Name.Text, sb.ToString()).PreserveSpace()));
                            sb = new StringBuilder();
                        }
                        newRuns.Add(new XElement(Name.Run, rPr, breakRun));
                        break;

                    default:
                        sb.Append(c);
                        break;
                }
            }

            if (sb.Length > 0)
            {
                newRuns.Add(new XElement(Name.Run, rPr,
                    new XElement(Name.Text, sb.ToString()).PreserveSpace()));
            }

            return newRuns;
        }

        internal static XElement[] SplitParagraph(Paragraph p, int index)
        {
            if (p == null)
                throw new ArgumentNullException(nameof(p));
            
            var r = p.GetFirstRunAffectedByEdit(index);
            XElement[] split;
            XElement before, after;

            switch (r.Xml.Parent?.Name.LocalName)
            {
                case "ins":
                    split = p.SplitEdit(r.Xml.Parent, index, EditType.Insert);
                    before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                    after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
                    break;
                case "del":
                    split = p.SplitEdit(r.Xml.Parent, index, EditType.Delete);
                    before = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split[0]);
                    after = new XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split[1]);
                    break;
                default:
                    split = r.SplitRun(index);
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

        /// <summary>
        /// Generate a 4-digit identifier and return the hex
        /// representation of it. 
        /// </summary>
        /// <param name="zeroPrefix">Number of bytes to be zero</param>
        internal static string GenerateHexId(int zeroPrefix = 0)
        {
            var data = new byte[4];
            new Random().NextBytes(data);

            for (int i = 0; i < Math.Min(zeroPrefix, data.Length); i++)
                data[i] = 0;

            return BitConverter.ToString(data).Replace("-", "");
        }

        /// <summary>
        /// Generate a 4-byte revision stamp from the current time.
        /// </summary>
        /// <returns>New revision stamp</returns>
        public static string GenerateRevisionStamp(string lastRevision, out uint newValue)
        {
            lastRevision ??= GenerateHexId(2);

            var dt = DateTime.Now;
            var val = uint.Parse(lastRevision, NumberStyles.AllowHexSpecifier);
            newValue = val + (uint) (dt.Second + dt.Millisecond);
            
            return val.ToString("X8");
        }
    }
}