using DXPlus.Helpers;
using System.Collections;
using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// Base class for header/footer collection
/// </summary>
/// <typeparam name="T"></typeparam>
public abstract class HeaderOrFooterCollection<T> : IEnumerable<T>
    where T : HeaderOrFooter, new()
{
    private readonly Section sectionOwner;
    private readonly Document documentOwner;
    private readonly string rootElementName;
    private readonly Relationship relationTemplate;
    private readonly string typeName;

    /// <summary>
    /// Constructor used to create the header collection
    /// </summary>
    /// <param name="documentOwner">Document owner</param>
    /// <param name="sectionOwner">Section owner</param>
    /// <param name="rootElementName">Root XML element name</param>
    /// <param name="relation">Relation name</param>
    /// <param name="typeName">Type name</param>
    internal HeaderOrFooterCollection(Document documentOwner, Section sectionOwner,
        string rootElementName, Relationship relation, string typeName)
    {
        this.documentOwner = documentOwner;
        this.sectionOwner = sectionOwner;
        this.rootElementName = rootElementName;
        relationTemplate = relation;
        this.typeName = typeName;

        First = LoadFromPackage(HeaderFooterType.First);
        Even = LoadFromPackage(HeaderFooterType.Even);
        Default = LoadFromPackage(HeaderFooterType.Odd);
    }

    /// <summary>
    /// Header/Footer on first page
    /// </summary>
    public T First { get; }

    /// <summary>
    /// Header/Footer on even page
    /// </summary>
    public T Even { get; }

    /// <summary>
    /// Header/Footer on odd pages
    /// </summary>
    public T Default { get; }

    /// <summary>
    /// Retrieve a header/footer from the owning document by type
    /// </summary>
    /// <param name="headerType">Header/Footer type (even, odd, default)</param>
    private T LoadFromPackage(HeaderFooterType headerType)
    {
        var id = GetReferenceId(headerType);
        if (id == null)
        {
            // Does not exist yet in the underlying document/package.
            // Leave that null for now - when this is added to the document
            // we can set the owners.
            return new T
            {
                Type = headerType,
                CreateFunc = Create,
                DeleteFunc = Delete,
                ExistsFunc = Exists
            };
        }

        // Load the header/footer
        documentOwner.FindHeaderFooterById(id, out var part, out var doc);

        var element = doc.Element(Namespace.Main + rootElementName);
        if (element == null) throw new DocumentFormatException(rootElementName);

        // Create the header/footer wrapper object from the loaded information
        var hf = new T
        {
            Xml = element,
            Id = id,
            Type = headerType,
            CreateFunc = Create,
            DeleteFunc = Delete,
            ExistsFunc = Exists
        };

        hf.SetOwner(documentOwner, part, false);
        return hf;
    }

    /// <summary>
    /// Looks up a header/footer using the ID and returns whether the relationship still exists.
    /// </summary>
    /// <param name="id">Header or footer id</param>
    /// <returns>True/False</returns>
    private bool Exists(string id) 
        => documentOwner.PackagePart.RelationshipExists(id);

    /// <summary>
    /// Adds a Header or Footer to a document.
    /// If the document already contains a Header it will be replaced.
    /// </summary>
    private void Create(HeaderOrFooter headerFooter)
    {
        var headerType = headerFooter.Type;

        // Delete any existing object; we have an id on the passed object
        if (headerFooter.Id != null)
        {
            Delete(headerFooter.Id, headerFooter.Type);
        }

        // Get the next header/footer index.
        var relations = documentOwner.PackagePart.GetRelationships()
            .Where(rel => rel.RelationshipType == relationTemplate.RelType)
            .Select(rel => rel.TargetUri.OriginalString)
            .Select(name => name.StartsWith("/word/") ? name : "/word/" + name)
            .Select(name => name.ToLower())
            .ToList();

        int index = 1;
        string filename = string.Format(relationTemplate.Path, index);
        while (relations.Contains(filename))
        {
            filename = string.Format(relationTemplate.Path, ++index);
        }

        var packagePart = documentOwner.Package.CreatePart(new Uri(filename, UriKind.Relative),
            relationTemplate.ContentType, CompressionOption.Normal);

        // Create the hdr/ftr definition
        var xmlFragment = XDocument.Parse(
            $@"<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>
                <w:{rootElementName} xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""{Namespace.RelatedDoc.NamespaceName}"" xmlns:m=""{Namespace.Math.NamespaceName}"" xmlns:v=""{Namespace.VML.NamespaceName}"" xmlns:wp=""{Namespace.WordProcessingDrawing.NamespaceName}"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""{Namespace.Main.NamespaceName}"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">
                    <w:p w:rsidR=""{documentOwner.RevisionId}"" w:rsidRDefault=""{documentOwner.RevisionId}"">
                        <w:pPr>
                            <w:pStyle w:val=""{typeName.ToCamelCase()}"" />
                        </w:pPr>
                    </w:p>
                </w:{rootElementName}>"
        );

        // Save the main document
        packagePart.Save(xmlFragment);

        // Add the relationship to the newly created header/footer
        var relationship = documentOwner.PackagePart.CreateRelationship(packagePart.Uri, TargetMode.Internal, relationTemplate.RelType);

        // Add the relationship to the owning section.
        sectionOwner.Properties.Xml.Add(new XElement(Namespace.Main + $"{typeName}Reference",
            new XAttribute(Namespace.Main + "type", headerType.GetEnumName()),
            new XAttribute(Namespace.RelatedDoc + "id", relationship.Id)));

        documentOwner.Package.Flush();

        // Let the document cache off the document.
        documentOwner.AdjustHeaderFooterCache(relationship.Id, xmlFragment);

        // Fill in the details.
        headerFooter.Xml = xmlFragment.Root!;
        headerFooter.Id = relationship.Id;
        headerFooter.SetOwner(documentOwner, packagePart, false);

        // If this is the first page header, then set the document.titlePg element
        if (headerFooter == First)
        {
            sectionOwner.Properties.DifferentFirstPage = true;
        }
        // Do the same for even page header
        else if (headerFooter == Even)
        {
            documentOwner.DifferentEvenOddHeadersFooters = true;
        }
    }

    /// <summary>
    /// Look up the header/footer reference in the main document
    /// </summary>
    /// <param name="headerType"></param>
    private string? GetReferenceId(HeaderFooterType headerType)
    {
        return sectionOwner.Properties.Xml
            .QueryElement($"w:{typeName}Reference[@w:type='{headerType.GetEnumName()}']")
            .AttributeValue(Namespace.RelatedDoc + "id", null);
    }

    /// <summary>
    /// Delete the specified header or footer
    /// </summary>
    private void Delete(HeaderOrFooter headerFooter)
    {
        if (headerFooter.Id != null)
        {
            Delete(headerFooter.Id, headerFooter.Type);
        }

        // If this is the first page header, then remove the document.titlePg element
        if (headerFooter == First)
        {
            sectionOwner.Properties.DifferentFirstPage = false;
        }
        // Do the same for even page header
        else if (headerFooter == Even)
        {
            documentOwner.DifferentEvenOddHeadersFooters = false;
        }
    }

    /// <summary>
    /// Delete the specified header or footer
    /// </summary>
    /// <param name="id"></param>
    /// <param name="headerType"></param>
    private void Delete(string id, HeaderFooterType headerType)
    {
        // Get this relationship.
        PackageRelationship rel = documentOwner.PackagePart.GetRelationship(id);
        Uri uri = rel.TargetUri;
        if (!uri.OriginalString.StartsWith("/word/"))
        {
            uri = new Uri("/word/" + uri.OriginalString, UriKind.Relative);
        }

        // Check to see if the document actually contains the Part.
        if (documentOwner.Package.PartExists(uri))
        {
            // Delete the part
            documentOwner.Package.DeletePart(uri);

            // Get all references to this relationship in the document and remove them.
            sectionOwner.Properties.Xml
                .QueryElement($"w:{typeName}Reference[@w:type='{headerType.GetEnumName()}' and @r:id='{id}']")?
                .Remove();
        }

        // Delete the Relationship.
        documentOwner.Package.DeleteRelationship(rel.Id);
        documentOwner.AdjustHeaderFooterCache(rel.Id, null);
    }

    /// <summary>
    /// Save the collection to the specific documents.
    /// </summary>
    internal void Save()
    {
        Default.Save();
        First.Save();
        Even.Save();
    }

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<T> GetEnumerator()
    {
        if (Default.Exists)
        {
            yield return Default;
        }

        if (First.Exists)
        {
            yield return First;
        }

        if (Even.Exists)
        {
            yield return Even;
        }
    }

    /// <summary>
    /// Returns the enumerator for the header/footers.
    /// </summary>
    /// <returns></returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}