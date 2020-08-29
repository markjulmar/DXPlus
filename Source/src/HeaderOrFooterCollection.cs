using System;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Base class for header/footer collection
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class HeaderOrFooterCollection<T> where T : HeaderOrFooter, new()
    {
        private readonly DocX documentOwner;
        private readonly string rootElementName;
        private readonly Relationship relationTemplate;
        private readonly string typeName;

        /// <summary>
        /// Constructor used to create the header collection
        /// </summary>
        /// <param name="documentOwner"></param>
        /// <param name="rootElementName"></param>
        /// <param name="relation"></param>
        /// <param name="typeName"></param>
        internal HeaderOrFooterCollection(DocX documentOwner, string rootElementName, Relationship relation, string typeName)
        {
            this.documentOwner = documentOwner;
            this.rootElementName = rootElementName;
            this.relationTemplate = relation;
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
        protected T LoadFromPackage(HeaderFooterType headerType)
        {
            string id = GetReferenceId(headerType);
            if (id == null)
            {
                return new T() {
                    Type = headerType,
                    CreateFunc = this.Create,
                    DeleteFunc = this.Delete
                };
            }

            // Get the Xml file for this Header or Footer.
            Uri partUri = documentOwner.PackagePart.GetRelationship(id).TargetUri;
            if (!partUri.OriginalString.StartsWith("/word/"))
                partUri = new Uri("/word/" + partUri.OriginalString, UriKind.Relative);

            // Get the PackagePart and load the XML
            var part = documentOwner.Package.GetPart(partUri);
            var doc = part.Load();

            // Create the header/footer from the loaded information
            return new T {
                Document = documentOwner,
                Xml = doc.Element(Namespace.Main + rootElementName),
                Id = id,
                Type = headerType,
                PackagePart = part,
                CreateFunc = this.Create,
                DeleteFunc = this.Delete
            };
        }

        /// <summary>
        /// Adds a Header or Footer to a document.
        /// If the document already contains a Header it will be replaced.
        /// </summary>
        internal void Create(HeaderOrFooter headerFooter)
        {
            var headerType = headerFooter.Type;

            // Delete any existing object; we have an id on the passed object
            if (headerFooter.Id != null)
            {
                Delete(headerFooter.Id, headerFooter.Type);
            }

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

            // Save the main document
            packagePart.Save(xmlFragment);

            // Add the relationship to the newly created header/footer
            var relationship = documentOwner.PackagePart.CreateRelationship(packagePart.Uri, TargetMode.Internal, relationTemplate.RelType);

            documentOwner.Xml.GetOrCreateElement(Name.SectionProperties)
                             .Add(new XElement(Namespace.Main + $"{typeName}Reference",
                                    new XAttribute(Namespace.Main + "type", headerType.GetEnumName()),
                                    new XAttribute(Namespace.RelatedDoc + "id", relationship.Id)));

            documentOwner.Package.Flush();

            // Fill in the details.
            headerFooter.Document = documentOwner;
            headerFooter.Xml = xmlFragment.Root;
            headerFooter.Id = relationship.Id;
            headerFooter.PackagePart = packagePart;

            // If this is the first page header, then set the document.titlePg element
            if (headerFooter == First)
            {
                documentOwner.DifferentFirstPage = true;
            }
            // Do the same for even page header
            else if (headerFooter == Even)
            {
                documentOwner.DifferentOddAndEvenPages = true;
            }
        }

        /// <summary>
        /// Look up the header/footer reference in the main document
        /// </summary>
        /// <param name="headerType"></param>
        private string GetReferenceId(HeaderFooterType headerType)
        {
            return documentOwner.QueryDocument($@"//w:sectPr/w:{typeName}Reference[@w:type='{headerType.GetEnumName()}']")
                .Select(e => e.AttributeValue(Namespace.RelatedDoc + "id"))
                .FirstOrDefault();
        }

        /// <summary>
        /// Delete the specified header or footer
        /// </summary>
        internal void Delete(HeaderOrFooter headerFooter)
        {
            if (headerFooter.Id != null)
            {
                Delete(headerFooter.Id, headerFooter.Type);
            }

            // If this is the first page header, then remove the document.titlePg element
            if (headerFooter == First)
            {
                documentOwner.DifferentFirstPage = false;
            }
            // Do the same for even page header
            else if (headerFooter == Even)
            {
                documentOwner.DifferentOddAndEvenPages = false;
            }
        }

        /// <summary>
        /// Delete the specified header or footer
        /// </summary>
        /// <param name="id"></param>
        private void Delete(string id, HeaderFooterType headerType)
        {
            // Get this relationship.
            var rel = documentOwner.PackagePart.GetRelationship(id);
            var uri = rel.TargetUri;
            if (!uri.OriginalString.StartsWith("/word/"))
                uri = new Uri("/word/" + uri.OriginalString, UriKind.Relative);

            // Check to see if the document actually contains the Part.
            if (documentOwner.Package.PartExists(uri))
            {
                // Delete the part
                documentOwner.Package.DeletePart(uri);

                // Get all references to this relationship in the document and remove them.
                documentOwner.QueryDocument($@"//w:sectPr/w:{typeName}Reference[@w:type='{headerType.GetEnumName()}' and @r:id='{id}']")
                             .SingleOrDefault()?.Remove();
            }

            // Delete the Relationship.
            documentOwner.Package.DeleteRelationship(rel.Id);
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
    }
}