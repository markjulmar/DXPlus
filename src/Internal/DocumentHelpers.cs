using System.Globalization;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DXPlus.Internal;

/// <summary>
/// Internal helper classes.
/// </summary>
internal static class DocumentHelpers
{
    /// <summary>
    /// Returns the Uri associated with a drawing/picture hyperlink. 
    /// </summary>
    /// <param name="xmlOwner">XML element (drawing or picture)</param>
    /// <param name="part">Package the element is in</param>
    /// <returns>Uri</returns>
    internal static Uri? GetHlinkClick(XElement xmlOwner, PackagePart part)
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
    /// Method to persist XML back to a PackagePart
    /// </summary>
    /// <param name="packagePart"></param>
    /// <param name="document"></param>
    internal static void Save(this PackagePart packagePart, XDocument document)
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
    internal static XDocument Load(this PackagePart packagePart)
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
    internal static bool IsValidHexNumber(string value)
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
    internal static int GetSize(XElement xml)
    {
        if (xml == null) throw new ArgumentNullException(nameof(xml));

        switch (xml.Name.LocalName)
        {
            case RunTextType.Tab:
            case RunTextType.CarriageReturn:
            case RunTextType.LineBreak:
                return 1;

            case RunTextType.Text:
            case RunTextType.DeletedText:
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
    internal static int? GetId(XElement element) 
        => int.TryParse(element.Attribute(Name.Id)?.Value ?? "", out var value) ? value : null;

    /// <summary>
    /// Return the inner text for a document element.
    /// </summary>
    /// <param name="element">XML element</param>
    /// <returns>Text or empty string</returns>
    internal static string GetText(XElement element) 
        => GetTextRecursive(element).ToString();

    /// <summary>
    /// Helper to enumerate through a document element and get all the text
    /// contained within as a single string.
    /// </summary>
    /// <param name="xml">Document element (Paragraph, Run, Hyperlink, etc.)</param>
    /// <param name="sb">StringBuilder used to build up text (pass null to start)</param>
    /// <returns>Text</returns>
    /// <exception cref="ArgumentNullException"></exception>
    internal static StringBuilder GetTextRecursive(XElement xml, StringBuilder? sb = null)
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
    internal static string ToText(XElement e)
    {
        switch (e.Name)
        {
            case {LocalName: RunTextType.Tab}:
                return "\t";

            case {LocalName: RunTextType.CarriageReturn}:
            case {LocalName: RunTextType.LineBreak}:
                return "\n";

            case {LocalName: RunTextType.Text}:
            case {LocalName: RunTextType.DeletedText}:
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
    /// Converts a string into a set of document Run objects with optional formatting.
    /// </summary>
    /// <param name="text">Text to parse</param>
    /// <param name="rPr">Optional run properties</param>
    /// <returns>List of run objects</returns>
    internal static IEnumerable<XElement> FormatInput(string text, XElement? rPr)
    {
        var newRuns = new List<XElement>();
        if (string.IsNullOrEmpty(text))
            return newRuns;

        var tab = new XElement(Namespace.Main + RunTextType.Tab);
        var lineBreak = new XElement(Namespace.Main + RunTextType.LineBreak);
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
                    newRuns.Add(new XElement(Name.Run, rPr, tab.Clone()));
                    break;

                case '\r':
                case '\n':
                    if (sb.Length > 0)
                    {
                        newRuns.Add(new XElement(Name.Run, rPr,
                            new XElement(Name.Text, sb.ToString()).PreserveSpace()));
                        sb = new StringBuilder();
                    }
                    newRuns.Add(new XElement(Name.Run, rPr, lineBreak.Clone()));
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
    /// Generate a 4-digit identifier and return the hex
    /// representation of it.
    /// </summary>
    /// <param name="zeroPrefix">Number of bytes to be zero</param>
    internal static string GenerateHexId(int zeroPrefix = 0)
    {
        byte[] data = new byte[4];
        new Random().NextBytes(data);

        for (int i = 0; i < Math.Min(zeroPrefix, data.Length); i++)
            data[i] = 0;

        return BitConverter.ToString(data).Replace("-", "");
    }

    /// <summary>
    /// Get the rPr element from a parent.
    /// </summary>
    /// <param name="owner"></param>
    /// <returns></returns>
    internal static XElement? GetRunProperties(this XElement owner) 
        => owner.Element(Name.RunProperties);

    /// <summary>
    /// Get or create the rPr element from a parent.
    /// </summary>
    /// <param name="owner"></param>
    /// <returns></returns>
    internal static XElement CreateRunProperties(this XElement owner)
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
    internal static XElement PreserveSpace(this XElement e)
    {
        if (!e.Name.Equals(Name.Text)
            && !e.Name.Equals(Namespace.Main + RunTextType.DeletedText))
        {
            throw new ArgumentException($"{nameof(PreserveSpace)} can only work with elements of type '{RunTextType.Text}' or '{RunTextType.DeletedText}'", nameof(e));
        }

        // Check if this w:t contains a space attribute
        var space = e.Attributes().SingleOrDefault(a => a.Name.Equals(XNamespace.Xml + "space"));

        // This w:t's text begins or ends with whitespace
        if (e.Value.StartsWith(' ') || e.Value.EndsWith(' '))
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
    /// Finds the next free Id for bookmarkStart/docPr.
    /// </summary>
    internal static long FindLastUsedDocId(XDocument document)
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
    /// Create a page break paragraph element
    /// </summary>
    /// <returns>Page break element</returns>
    internal static XElement PageBreak() => new(Name.Paragraph,
        new XAttribute(Name.ParagraphId, GenerateHexId()),
        new XElement(Name.Run, new XElement(Namespace.Main + RunTextType.LineBreak,
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
                    case RunTextType.Tab:
                        if (el.Parent?.Name.LocalName != "tabs") count++;
                        break;

                    case RunTextType.LineBreak:
                        count++;
                        break;

                    case RunTextType.Text:
                    case RunTextType.DeletedText:
                        count += el.Value.Length;
                        break;
                }
            }
        }
        return count;
    }
}