﻿using DXPlus.Resources;
using System.Diagnostics;
using System.Globalization;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DXPlus.Helpers;

/// <summary>
/// Internal helper classes.
/// </summary>
internal static class HelperFunctions
{
    /// <summary>
    /// Returns the Uri associated with a drawing/picture hyperlink. 
    /// </summary>
    /// <param name="xmlOwner">XML element (drawing or picture)</param>
    /// <param name="part">Package the element is in</param>
    /// <returns>Uri</returns>
    public static Uri? GetHlinkClick(XElement xmlOwner, PackagePart part)
    {
        if (xmlOwner == null) throw new ArgumentNullException(nameof(xmlOwner));
        if (part == null) throw new ArgumentNullException(nameof(part));
        string? id = xmlOwner.Element(Namespace.DrawingMain + "hlinkClick")
                             ?.AttributeValue(Namespace.RelatedDoc + "id");
        return string.IsNullOrEmpty(id) ? null : part.GetRelationship(id).TargetUri;
    }

    /// <summary>
    /// Set the Uri associated with a drawing/picture hyperlink
    /// </summary>
    /// <param name="xmlOwner">XML element (drawing or picture)</param>
    /// <param name="part">Package the element is in</param>
    /// <param name="uri">Uri to assign</param>
    internal static void SetHlinkClick(XElement xmlOwner, PackagePart part, Uri? uri)
    {
        if (xmlOwner == null) throw new ArgumentNullException(nameof(xmlOwner));
        if (part == null) throw new ArgumentNullException(nameof(part));
        if (uri == null)
        {
            // Delete the relationship
            string? id = xmlOwner.Element(Namespace.DrawingMain + "hlinkClick")
                                ?.AttributeValue(Namespace.RelatedDoc + "id");
            if (!string.IsNullOrEmpty(id)) part.DeleteRelationship(id);
            xmlOwner.Element(Namespace.DrawingMain + "hlinkClick")?.Remove();
            return;
        }

        // Create a new relationship.
        TargetMode targetMode = TargetMode.External;
        string relationshipType = $"{Namespace.RelatedDoc.NamespaceName}/hyperlink";
        string url = uri.OriginalString;
        string rid = part.GetRelationshipsByType(relationshipType)
                         .Where(r => r.TargetUri.OriginalString == url)
                         .Select(r => r.Id)
                         .SingleOrDefault() ??
                     part.CreateRelationship(uri, targetMode, relationshipType).Id;

        part.DeleteRelationship(rid);
        part.CreateRelationship(uri, targetMode, relationshipType, rid);
                
        // Add the hyperlink to the XML.
        xmlOwner.Element(Namespace.DrawingMain + "hlinkClick")?.Remove();
        xmlOwner.Add(new XElement(Namespace.DrawingMain + "hlinkClick",
            new XAttribute(Namespace.RelatedDoc + "id", rid)));
    }
     
    /// <summary>
    /// Wraps a block container from the document. These are elements
    /// which contain other elements.
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package part</param>
    /// <param name="e">Element</param>
    /// <returns></returns>
    internal static BlockContainer? WrapElementBlockContainer(Document? document, PackagePart? packagePart, XElement e)
    {
        if (e.Name == Name.TableCell)
        {
            if (e.Parent?.Parent != null)
            {
                var rowXml = e.Parent;
                var tableXml = rowXml.Parent;
                var table = new Table(document, packagePart, tableXml);
                var row = new TableRow(table, rowXml);
                return new TableCell(row, e);
            }
        }

        if (e.Name == Name.Body)
        {
            return document;
        }

        if (e.Name.LocalName == "hdr" && document != null)
        {
            return document.Sections.SelectMany(s => s.Headers)
                                    .SingleOrDefault(h => h.Xml == e);
        }

        if (e.Name.LocalName == "ftr" && document != null)
        {
            return document.Sections.SelectMany(s => s.Footers)
                                    .SingleOrDefault(h => h.Xml == e);
        }

        throw new Exception($"Unrecognized container type {e.Name}");
    }

    /// <summary>
    /// Wrap an element
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package part owner</param>
    /// <param name="xml">XML fragment</param>
    /// <returns>Element wrapper</returns>
    internal static DocXElement? WrapDocumentElement(Document document, PackagePart packagePart, XElement xml)
    {
        if (xml.Name.LocalName == "hyperlink")
            return new Hyperlink(document, packagePart, xml);

        if (xml.Name == Name.Paragraph)
        {
            // See if we can find it first. That way we get the proper index.
            var p = document.Paragraphs.FirstOrDefault(p => p.Xml == xml);
            if (p != null) return p;

            // Hmm. Unowned perhaps?
            int pos = 0;
            return WrapParagraphElement(xml, document, packagePart, ref pos);
        }
        return null;
    }

    /// <summary>
    /// Wraps a paragraph object
    /// </summary>
    /// <param name="element">XML element</param>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package paragraph is in</param>
    /// <param name="position">Text position</param>
    /// <returns>Paragraph wrapper</returns>
    internal static Paragraph WrapParagraphElement(XElement element, Document? document, PackagePart? packagePart, ref int position)
    {
        if (element.Name != Name.Paragraph)
            throw new ArgumentException($"Passed element {element.Name} not a {Name.Paragraph}.", nameof(element));

        var p = new Paragraph(document, packagePart, element, position);
        position += GetText(element).Length;
        return p;
    }

    /// <summary>
    /// Helper to create a block from an element in the document.
    /// </summary>
    /// <param name="blockContainer">Owning container</param>
    /// <param name="e">XML element</param>
    /// <param name="current">Current text position for paragraph tracking</param>
    /// <returns>Block wrapper</returns>
    internal static Block? WrapElementBlock(BlockContainer blockContainer, XElement e, ref int current)
    {
        if (blockContainer == null) throw new ArgumentNullException(nameof(blockContainer));
        if (e == null) throw new ArgumentNullException(nameof(e));

        if (e.Name == Name.Paragraph)
        {
            return WrapParagraphElement(e, blockContainer.Document, blockContainer.PackagePart, ref current);
        }
        if (e.Name == Name.Table)
        {
            return new Table(blockContainer.Document, blockContainer.PackagePart, e);
        }
        if (e.Name != Name.SectionProperties)
        {
            return new UnknownBlock(blockContainer.Document, blockContainer.PackagePart, e);
        }
        return null;
    }

    /// <summary>
    /// This creates a Word docx in a memory stream.
    /// </summary>
    /// <param name="documentType">Type (doc or template)</param>
    /// <returns>Memory stream with loaded doc</returns>
    public static Stream CreateDocumentType(DocumentTypes documentType)
    {
        // Create the docx package
        var ms = new MemoryStream();
        using var package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);

        // Force app/xml to be registered as the default document type
        var appPath = new Uri($"/app.xml", UriKind.Relative);
        _ = package.CreatePart(appPath, "application/xml");

        // Create the main document part for this package
        var mainDocumentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative),
            documentType == DocumentTypes.Document ? DocxContentType.Document : DocxContentType.Template,
            CompressionOption.Normal);
        package.CreateRelationship(mainDocumentPart.Uri, TargetMode.Internal, $"{Namespace.RelatedDoc.NamespaceName}/officeDocument");

        // We don't actually need a real file -- just the <Default/> tag.
        package.DeletePart(appPath);

        // Generate an id for this editing session.
        string startingRevisionId = GenerateRevisionStamp(null, out _);

        // Load the document part into a XDocument object
        XDocument mainDoc = Resource.BodyDocument(startingRevisionId);

        // Add the settings.xml + relationship
        _ = AddDefaultSettingsPart(package, startingRevisionId);

        // Add the default styles + relationship
        _ = AddDefaultStylesXml(package, out _);

        // Create the document properties.
        _ = CorePropertyHelpers.CreateCoreProperties(package, out _);

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
    public static void Save(this PackagePart packagePart, XDocument document)
    {
        if (packagePart == null)
        {
            throw new ArgumentNullException(nameof(packagePart));
        }

        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        lock (packagePart.Package)
        {
            using var writer = new StreamWriter(packagePart.GetStream(FileMode.Create, FileAccess.Write));
            var xmlWriter = XmlWriter.Create(writer, new XmlWriterSettings
            {
                Indent = false,
                CheckCharacters = true,
                CloseOutput = true,
                ConformanceLevel = ConformanceLevel.Document,
                Encoding = Encoding.UTF8,
                NamespaceHandling = NamespaceHandling.OmitDuplicates,
                NewLineHandling = NewLineHandling.None,
                OmitXmlDeclaration = false
            });
            document.Save(xmlWriter);
            xmlWriter.Close();
        }

#if DEBUG
        //_ = Load(packagePart);
#endif
    }

    /// <summary>
    /// Load XML from a package part
    /// </summary>
    /// <param name="packagePart"></param>
    /// <returns></returns>
    public static XDocument Load(this PackagePart packagePart)
    {
        if (packagePart == null)
        {
            throw new ArgumentNullException(nameof(packagePart));
        }

        lock (packagePart.Package)
        {
            using var reader = new StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8);
            var document = XDocument.Load(reader, LoadOptions.None);
            if (document.Root == null)
            {
                throw new Exception("Loaded document from PackagePart has no contents.");
            }
            return document;
        }
    }

    /// <summary>
    /// Helper to check a string and verify it's a valid HEX value.
    /// </summary>
    /// <param name="value">String to check</param>
    /// <returns>True/False if the value is a hex string</returns>
    public static bool IsValidHexNumber(string value)
    {
        return !string.IsNullOrWhiteSpace(value)
               && value.Length <= 8
               && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int _);
    }

    /// <summary>
    /// Helper to return the text length of a document element.
    /// </summary>
    /// <param name="xml">Document element</param>
    /// <returns>Length</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static int GetSize(XElement xml)
    {
        if (xml == null) throw new ArgumentNullException(nameof(xml));

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

    /// <summary>
    /// Return the unique identifier assigned to a document element
    /// </summary>
    /// <param name="element">XML element</param>
    /// <returns>Integer identifier if present</returns>
    public static int? GetId(XElement element) 
        => int.TryParse(element.Attribute(Name.Id)?.Value ?? "", out var value) ? value : null;

    /// <summary>
    /// Return the inner text for a document element.
    /// </summary>
    /// <param name="element">XML element</param>
    /// <returns>Text or empty string</returns>
    public static string GetText(XElement element) 
        => GetTextRecursive(element).ToString();

    /// <summary>
    /// Helper to enumerate through a document element and get all the text
    /// contained within as a single string.
    /// </summary>
    /// <param name="xml">Document element (Paragraph, Run, Hyperlink, etc.)</param>
    /// <param name="sb">StringBuilder used to build up text (pass null to start)</param>
    /// <returns>Text</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static StringBuilder GetTextRecursive(XElement xml, StringBuilder? sb = null)
    {
        if (xml == null) throw new ArgumentNullException(nameof(xml));

        sb ??= new StringBuilder();

        // Skip property blocks.
        if (xml.Name == Name.ParagraphProperties
            || xml.Name == Name.RunProperties)
            return sb;
                
        string text = ToText(xml);
        if (!string.IsNullOrEmpty(text))
            sb.Append(text);

        if (xml.HasElements)
            sb = xml.Elements().Aggregate(sb, (current, e) => GetTextRecursive(e, current));

        return sb;
    }

    /// <summary>
    /// Turn a Word (w:t) element into text.
    /// </summary>
    /// <param name="e"></param>
    public static string ToText(XElement e)
    {
        switch (e.Name)
        {
            case {LocalName: "tab"}:
            case {LocalName: "tc"}:
                return "\t";

            case {LocalName: "cr"}:
            case {LocalName: "tr"}:
            case {LocalName: "br"}:
                return "\n";

            case {LocalName: "t"}:
            case {LocalName: "delText"}:
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
    public static XElement Clone(this XElement element)
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
    public static XDocument AddDefaultStylesXml(Package package, out PackagePart stylesPart)
    {
        if (package == null) throw new ArgumentNullException(nameof(package));

        if (package.PartExists(Relations.Styles.Uri))
            throw new InvalidOperationException("Root style collection already exists.");

        stylesPart = package.CreatePart(Relations.Styles.Uri, Relations.Styles.ContentType, CompressionOption.Maximum);
        var stylesDoc = Resource.DefaultStylesXml();

        Debug.Assert(stylesDoc.Root != null);
        Debug.Assert(stylesDoc.Root.Element(Namespace.Main + "docDefaults") != null);

        // Set the run default language to be the current culture.
        stylesDoc.Root.Element(Namespace.Main + "docDefaults",
                Namespace.Main + "rPrDefault",
                Name.RunProperties, Name.Language)!
            .SetAttributeValue(Name.MainVal, CultureInfo.CurrentCulture);

        // Save /word/styles.xml
        stylesPart.Save(stylesDoc);

        // Add the relationship to the main doc
        PackagePart mainDocumentPart = package.GetParts().Single(p =>
            p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
            || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

        mainDocumentPart.CreateRelationship(stylesPart.Uri, TargetMode.Internal, $"{Namespace.RelatedDoc.NamespaceName}/styles");
            
        return stylesDoc;
    }

    /// <summary>
    /// Create the /word/settings.xml document
    /// </summary>
    /// <param name="package">Package owner</param>
    /// <param name="rsid">Initial document revision id</param>
    public static XDocument AddDefaultSettingsPart(Package package, string rsid)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        if (string.IsNullOrEmpty(rsid))
            throw new ArgumentException($"'{nameof(rsid)}' cannot be null or empty", nameof(rsid));
        if (package.PartExists(Relations.Settings.Uri))
            throw new InvalidOperationException("Settings.xml section already exists.");

        // Add the settings package part and document
        PackagePart settingsPart = package.CreatePart(Relations.Settings.Uri, Relations.Settings.ContentType, CompressionOption.Maximum);
        var docId = Guid.NewGuid().ToString();
        XDocument settings = Resource.SettingsXml(rsid, docId[..8], "{" + docId + "}");

        Debug.Assert(settings.Root != null);

        // Set the correct language
        settings.Root.Element(Namespace.Main + "themeFontLang")!
            .SetAttributeValue(Name.MainVal, CultureInfo.CurrentCulture);

        // Save the settings document.
        settingsPart.Save(settings);

        // Add the relationship to the main doc
        var mainDocumentPart = package.GetParts().Single(p =>
            p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase)
            || p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

        mainDocumentPart.CreateRelationship(Relations.Settings.Uri, TargetMode.Internal, Relations.Settings.RelType);

        return settings;
    }

    /// <summary>
    /// Converts a string into a set of document Run objects with optional formatting.
    /// </summary>
    /// <param name="text">Text to parse</param>
    /// <param name="rPr">Optional run properties</param>
    /// <returns>List of run objects</returns>
    public static List<XElement> FormatInput(string text, XElement? rPr)
    {
        var newRuns = new List<XElement>();
        if (string.IsNullOrEmpty(text))
            return newRuns;

        var tab = new XElement(Namespace.Main + "tab");
        var lineBreak = new XElement(Namespace.Main + "br");
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
                    newRuns.Add(new XElement(Name.Run, rPr, tab));
                    break;

                case '\r':
                case '\n':
                    if (sb.Length > 0)
                    {
                        newRuns.Add(new XElement(Name.Run, rPr,
                            new XElement(Name.Text, sb.ToString()).PreserveSpace()));
                        sb = new StringBuilder();
                    }
                    newRuns.Add(new XElement(Name.Run, rPr, lineBreak));
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

    /// <summary>
    /// Helper to determine if two streams represent the same physical file.
    /// </summary>
    /// <param name="streamOne">First stream</param>
    /// <param name="streamTwo">Second stream</param>
    /// <returns>True/False</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static bool IsSameFile(Stream streamOne, Stream streamTwo)
    {
        if (streamOne == null) throw new ArgumentNullException(nameof(streamOne));
        if (streamTwo == null) throw new ArgumentNullException(nameof(streamTwo));

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

    /// <summary>
    /// Generate a 4-digit identifier and return the hex
    /// representation of it.
    /// </summary>
    /// <param name="zeroPrefix">Number of bytes to be zero</param>
    public static string GenerateHexId(int zeroPrefix = 0)
    {
        byte[] data = new byte[4];
        new Random().NextBytes(data);

        for (int i = 0; i < Math.Min(zeroPrefix, data.Length); i++)
            data[i] = 0;

        return BitConverter.ToString(data).Replace("-", "");
    }

    /// <summary>
    /// Generate a 4-byte revision stamp from the current time.
    /// </summary>
    /// <returns>New revision stamp</returns>
    public static string GenerateRevisionStamp(string? lastRevision, out uint newValue)
    {
        lastRevision ??= GenerateHexId(2);

        DateTime dt = DateTime.Now;
        uint val = uint.Parse(lastRevision, NumberStyles.AllowHexSpecifier);
        newValue = val + (uint)(dt.Second + dt.Millisecond);

        return val.ToString("X8");
    }

    /// <summary>
    /// Get the rPr element from a parent.
    /// </summary>
    /// <param name="owner"></param>
    /// <returns></returns>
    public static XElement? GetRunProperties(this XElement owner) 
        => owner.Element(Name.RunProperties);

    /// <summary>
    /// Get or create the rPr element from a parent.
    /// </summary>
    /// <param name="owner"></param>
    /// <returns></returns>
    public static XElement CreateRunProperties(this XElement owner)
    {
        XElement? rPr = owner.Element(Name.RunProperties);
        if (rPr == null)
        {
            rPr = new XElement(Name.RunProperties);
            owner.AddFirst(rPr); // must always be first.
        }
        return rPr;
    }

    /// <summary>
    /// If a text element or delText element, starts or ends with a space,
    /// it must have the attribute space, otherwise it must not have it.
    /// </summary>
    /// <param name="e">The (t or delText) element check</param>
    public static XElement PreserveSpace(this XElement e)
    {
        if (!e.Name.Equals(Name.Text)
            && !e.Name.Equals(Namespace.Main + "delText"))
        {
            throw new ArgumentException($"{nameof(PreserveSpace)} can only work with elements of type 't' or 'delText'", nameof(e));
        }

        // Check if this w:t contains a space attribute
        var space = e.Attributes().SingleOrDefault(a => a.Name.Equals(XNamespace.Xml + "space"));

        // This w:t's text begins or ends with whitespace
        if (e.Value.StartsWith(" ") || e.Value.EndsWith(" "))
        {
            // If this w:t contains no space attribute, add one.
            if (space == null)
            {
                e.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
            }
        }

        // This w:t's text does not begin or end with a space
        else
        {
            // If this w:r contains a space attribute, remove it.
            space?.Remove();
        }

        return e;
    }

    /// <summary>
    /// Gets/Creates the section properties for an owner
    /// </summary>
    /// <param name="element">Element owner</param>
    /// <param name="create">True to create</param>
    /// <returns>Section properties object</returns>
    public static SectionProperties GetSectionProperties(this XElement element, bool create = false)
    {
        var sectPr = element.Element(Name.SectionProperties);
        if (create && sectPr == null)
        {
            sectPr = new XElement(Name.SectionProperties);
            element.Add(sectPr);
        }
        return new SectionProperties(sectPr);
    }

    /// <summary>
    /// Finds the next free Id for bookmarkStart/docPr.
    /// </summary>
    public static long FindLastUsedDocId(XDocument document)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));

        var ids = document.Root!.Descendants()
            .Where(e => e.Name == Name.BookmarkStart || e.Name == Name.DrawingProperties)
            .Select(e => e.Attributes().FirstOrDefault(a => a.Name.LocalName == "id"))
            .Where(id => id != null)
            .Select(id => long.TryParse(id!.Value, out var value) ? value : 0)
            .Where(id => id > 0)
            .ToHashSet();

        return ids.Count > 0 ? ids.Max() : 0;
    }

    /// <summary>
    /// This helper splits an XML name into a namespace + localName.
    /// </summary>
    /// <param name="name">Full name</param>
    /// <param name="ns">Returning namespace</param>
    /// <param name="localName">Returning localName</param>
    public static bool SplitXmlName(string name, out string ns, out string localName)
    {
        if (name.Contains(':'))
        {
            var parts = name.Split(':');
            ns = parts[0];
            localName = parts[1];
            return true;
        }

        ns = string.Empty;
        localName = name;
        return false;
    }

    /// <summary>
    /// Create a page break paragraph element
    /// </summary>
    /// <returns>Page break element</returns>
    public static XElement PageBreak() => new(Name.Paragraph,
        new XAttribute(Name.ParagraphId, GenerateHexId()),
        new XElement(Name.Run, new XElement(Namespace.Main + "br",
            new XAttribute(Namespace.Main + "type", "page"))));

    /// <summary>
    /// Retrieve the text length of the passed element
    /// </summary>
    /// <param name="textElement"></param>
    /// <returns></returns>
    internal static int GetTextLength(XElement? textElement)
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
                        {
                            goto case "br";
                        }

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
}