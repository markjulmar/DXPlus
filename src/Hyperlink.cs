using DXPlus.Helpers;
using System.Diagnostics;
using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// Represents a Hyperlink in a document.
/// </summary>
public class Hyperlink : DocXElement, IEquatable<Hyperlink>
{
    private Uri? uri;
    private string text;
    private readonly int type; // 0 = simple hyperlink, 1 = complex instrText
    private readonly XElement? instrText;
    private readonly List<XElement>? runs;

    /// <summary>
    /// Unique id for this hyperlink
    /// </summary>
    public string? Id { get; private set; }

    /// <summary>
    /// Remove the hyperlink from the owning paragraph.
    /// </summary>
    public void Remove()
    {
        if (Xml.Parent != null)
        {
            Xml.Remove();
        }
    }

    /// <summary>
    /// Change the Text of a Hyperlink.
    /// </summary>
    public string Text
    {
        get => text;

        set
        {
            var rPr = new XElement(Name.RunProperties,
                new XElement(Namespace.Main + "rStyle",
                    new XAttribute(Name.MainVal, "Hyperlink")));

            // Format and add the new text.
            var newRuns = HelperFunctions.FormatInput(value, rPr);
            if (type == 0)
            {
                Xml.Elements(Name.Run).Remove();
                Xml.Add(newRuns);
            }
            else
            {
                var separate = XElement.Parse(@"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='separate'/>
                    </w:r>");

                var end = XElement.Parse(@"
                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                        <w:fldChar w:fldCharType='end' />
                    </w:r>");

                runs!.Last().AddAfterSelf(separate, newRuns, end);
                runs!.ForEach(r => r.Remove());
            }

            text = value;
        }
    }

    /// <summary>
    /// Change the Uri of a Hyperlink.
    /// </summary>
    public Uri Uri
    {
        get => type == 0 && !string.IsNullOrEmpty(Id)
            ? PackagePart.GetRelationship(Id).TargetUri
            : uri!;

        private set
        {
            uri = value;

            if (type == 0)
            {
                if (!string.IsNullOrEmpty(Id))
                {
                    var packageRelation = PackagePart.GetRelationship(Id);

                    // Get all of the information about this relationship.
                    TargetMode targetMode = packageRelation.TargetMode;
                    string relationshipType = packageRelation.RelationshipType;
                    string id = packageRelation.Id;

                    PackagePart.DeleteRelationship(id);
                    PackagePart.CreateRelationship(value, targetMode, relationshipType, id);
                }
            }
            else
            {
                instrText!.Value = $"HYPERLINK \"{value}\"";
            }
        }
    }

    /// <summary>
    /// Create a new hyperlink to add to a document
    /// </summary>
    /// <param name="text">Link text</param>
    /// <param name="uri">Link URI</param>
    public Hyperlink(string text, Uri uri)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(text));

        type = 0;
        Uri = uri ?? throw new ArgumentNullException(nameof(uri));
        this.text = text;
        Id = string.Empty;

        base.Xml = new XElement(Namespace.Main + "hyperlink",
            new XAttribute(Namespace.RelatedDoc + "id", string.Empty),
            new XAttribute(Namespace.Main + "history", "1"),
            new XElement(Name.Run,
                new XElement(Name.RunProperties,
                    new XElement(Namespace.Main + "rStyle",
                        new XAttribute(Name.MainVal, "Hyperlink"))),
                new XElement(Name.Text, text))
        );
    }

    /// <summary>
    /// Internal constructor used when creating hyperlinks out of the document
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package owner</param>
    /// <param name="xml">XML fragment representing hyperlink</param>
    internal Hyperlink(Document document, PackagePart packagePart, XElement xml) : base(xml)
    {
        type = 0;
        Id = xml.AttributeValue(Namespace.RelatedDoc + "id");
        text = HelperFunctions.GetTextRecursive(xml).ToString();

        SetOwner(document, packagePart, false);
    }

    /// <summary>
    /// Internal constructor used when creating hyperlinks out of the document
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package owner</param>
    /// <param name="instrText">Text with field codes</param>
    /// <param name="runs">Text runs making up this hyperlink</param>
    private Hyperlink(Document document, PackagePart packagePart, XElement instrText, List<XElement> runs)
    {
        type = 1;
        this.instrText = instrText;
        this.runs = runs;

        SetOwner(document, packagePart, false);

        int start = instrText.Value.IndexOf("HYPERLINK \"", StringComparison.Ordinal) + "HYPERLINK \"".Length;
        int end = instrText.Value.IndexOf("\"", start, StringComparison.Ordinal);
        if (start != -1 && end != -1)
        {
            Uri = new Uri(instrText.Value[start..end], UriKind.Absolute);
            text = HelperFunctions.GetTextRecursive(new XElement(Namespace.Main + "temp", runs)).ToString();
        }
        else
        {
            text = string.Empty;
        }
    }

    /// <summary>
    /// Get or create a relationship link to a hyperlink
    /// </summary>
    /// <returns>Relationship id</returns>
    internal string GetOrCreateRelationship()
    {
        string baseUri = Uri.OriginalString;

        // Search for a relationship with a TargetUri for this hyperlink.
        var id = PackagePart.GetRelationshipsByType($"{Namespace.RelatedDoc.NamespaceName}/hyperlink")
            .Where(r => r.TargetUri.OriginalString == baseUri)
            .Select(r => r.Id)
            .SingleOrDefault();

        // No id yet, create one.
        if (id == null)
        {
            // Check to see if a relationship for this Hyperlink exists and create it if not.
            var packageRelation = PackagePart.CreateRelationship(Uri, TargetMode.External,
                $"{Namespace.RelatedDoc.NamespaceName}/hyperlink");
            id = packageRelation.Id;
        }

        Id = id;
        Xml.SetAttributeValue(Namespace.RelatedDoc + "id", id);

        return id;
    }


    /// <summary>
    /// Return all hyperlinks associated to a given parent
    /// </summary>
    /// <param name="owner"></param>
    /// <param name="unownedHyperlinks"></param>
    /// <returns></returns>
    internal static IEnumerable<Hyperlink> Enumerate(DocXElement owner, IList<Hyperlink> unownedHyperlinks)
    {
        var allHyperlinks = owner.Xml.Descendants()
            .Where(h => h.Name.LocalName is "hyperlink" or "instrText")
            .ToList();

        foreach (var he in allHyperlinks)
        {
            if (he.Name.LocalName == "hyperlink")
            {
                var hyperlink = unownedHyperlinks?.SingleOrDefault(hl => ReferenceEquals(hl.Xml, he));
                if (hyperlink != null)
                {
                    yield return hyperlink;
                    continue;
                }

                Debug.Assert(owner.Document != null);
                yield return new Hyperlink(owner.Document, owner.PackagePart, he);
            }
            else
            {
                // Find the parent run, no matter how deeply nested we are.
                var runParent = he.FindParent(Name.Run);
                if (runParent == null)
                    throw new Exception("Failed to locate the parent in a run.");

                // Take every element until we reach w:fldCharType="end"
                var hyperLinkRuns = new List<XElement>();
                foreach (var run in runParent.ElementsAfterSelf(Name.Run))
                {
                    // Add this run to the list.
                    hyperLinkRuns.Add(run);

                    var fldChar = run.Descendants(Namespace.Main + "fldChar").SingleOrDefault();
                    if (fldChar != null)
                    {
                        var fldCharType = fldChar.Attribute(Namespace.Main + "fldCharType");
                        if (fldCharType?.Value.Equals("end", StringComparison.CurrentCultureIgnoreCase) == true)
                        {
                            yield return new Hyperlink(owner.Document, owner.PackagePart, he, hyperLinkRuns);
                            break;
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Called when the document owner is changed.
    /// </summary>
    protected override void OnAddToDocument() => Document.AddHyperlinkStyle();

    /// <summary>
    /// Determines equality for hyperlinks
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Hyperlink? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);
}