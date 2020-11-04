using DXPlus.Charts;
using DXPlus.Helpers;
using DXPlus.Resources;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a document.
    /// </summary>
    internal sealed class DocX : BlockContainer, IDocument
    {
        private string filename;               // The filename that this document was loaded from; can be null;
        private Stream stream;                 // The stream that this document was loaded from; can be null.
        private MemoryStream memoryStream;     // The in-memory document (with changes)
        private long nextDocumentId;           // Next document id for pictures, bookmarks, etc.

        /// <summary>
        /// The ZIP package holding this DocX structure
        /// </summary>
        internal Package Package { get; set; }

        /// <summary>
        /// XML documents representing loaded sections of the DOCX file.
        /// These have possible unsaved edits
        /// </summary>
        private XDocument footnotesDoc;
        private XDocument endnotesDoc;
        private XDocument mainDoc;
        private XDocument settingsDoc;
        private readonly Dictionary<string, XDocument> headerFooterCache = new Dictionary<string, XDocument>();

        /// <summary>
        /// Package sections in the above Package object. These are specific read/write points in the ZIP file
        /// </summary>
        private PackagePart footnotesPart;
        private PackagePart endnotesPart;
        private PackagePart settingsPart;

        /// <summary>
        /// Document revision
        /// </summary>
        private uint revision;

        /// <summary>
        /// Default constructor
        /// </summary>
        internal DocX() : base(null, null)
        {
        }

        /// <summary>
        /// Editing session id for this session.
        /// </summary>
        public string RevisionId => revision.ToString("X8");

        /// <summary>
        /// Retrieve all bookmarks
        /// </summary>
        public BookmarkCollection Bookmarks => new BookmarkCollection(Paragraphs.SelectMany(p => p.GetBookmarks()));

        /// <summary>
        /// Numbering styles in this document.
        /// </summary>
        public NumberingStyleManager NumberingStyles { get; private set; }

        /// <summary>
        /// The style manager for this document.
        /// </summary>
        public StyleManager Styles { get; private set; }

        /// <summary>
        /// True if the Document use different Headers and Footers for odd and even pages.
        /// </summary>
        public bool DifferentEvenOddHeadersFooters
        {
            get
            {
                ThrowIfObjectDisposed();
                return settingsDoc.Root!.Element(Namespace.Main + "evenAndOddHeaders") != null;
            }

            set
            {
                ThrowIfObjectDisposed();
                var evenAndOddHeaders = settingsDoc.Root!.Element(Namespace.Main + "evenAndOddHeaders");
                if (evenAndOddHeaders == null && value)
                {
                    settingsDoc.Root.AddFirst(new XElement(Namespace.Main + "evenAndOddHeaders"));
                }
                else if (!value)
                {
                    evenAndOddHeaders?.Remove();
                }
            }
        }

        /// <summary>
        /// Get the text of each endnote from this document
        /// </summary>
        public IEnumerable<string> EndnotesText
            => endnotesDoc.Root?.Elements(Namespace.Main + "endnote").Select(HelperFunctions.GetText);

        /// <summary>
        /// Get the text of each footnote from this document
        /// </summary>
        public IEnumerable<string> FootnotesText
            => footnotesDoc.Root?.Elements(Namespace.Main + "footnote").Select(HelperFunctions.GetText);

        /// <summary>
        /// Returns a list of Images in this document.
        /// </summary>
        public List<Image> Images
        {
            get
            {
                ThrowIfObjectDisposed();
                var imageRelationships = PackagePart.GetRelationshipsByType($"{Namespace.RelatedDoc.NamespaceName}/image");
                return imageRelationships.Any()
                    ? imageRelationships.Select(i => new Image(this, i)).ToList()
                    : new List<Image>();
            }
        }

        /// <summary>
        /// Method to create an unnamed document
        /// </summary>
        /// <param name="filename">optional filename</param>
        /// <param name="documentType">Type to create</param>
        /// <returns>New document</returns>
        internal static DocX Create(string filename, DocumentTypes documentType)
        {
            var doc = Load(HelperFunctions.CreateDocumentType(documentType));
            doc.stream.Dispose();
            doc.stream = null;
            doc.filename = filename;
            doc.revision--;

            return doc;
        }

        /// <summary>
        /// Loads a document into a DocX object using a Stream.
        /// </summary>
        /// <param name="stream">The Stream to load the document from.</param>
        /// <returns>
        /// Returns a DocX object which represents the document.
        /// </returns>
        internal static DocX Load(Stream stream)
        {
            var ms = new MemoryStream();
            stream.Seek(0, SeekOrigin.Begin);
            stream.CopyTo(ms);

            // Open the docx package
            var package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

            var document = CreateDocumentFromPackage(package);
            document.memoryStream = ms;
            document.stream = stream;
            return document;
        }

        /// <summary>
        /// Loads a document into a DocX object using a fully qualified or relative filename.
        /// </summary>
        /// <param name="filename">The fully qualified or relative filename.</param>
        /// <returns>
        /// Returns a DocX object which represents the document.
        /// </returns>
        internal static DocX Load(string filename)
        {
            if (string.IsNullOrWhiteSpace(filename))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(filename));
            if (!File.Exists(filename))
                throw new FileNotFoundException("Docx file doesn't exist", filename);

            // Open the docx package
            var ms = new MemoryStream();
            using (var fs = new FileStream(filename, FileMode.Open))
            {
                fs.CopyTo(ms);
            }

            var package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

            var document = CreateDocumentFromPackage(package);
            document.filename = filename;
            document.memoryStream = ms;

            return document;
        }

        ///<summary>
        /// Returns the list of document properties with corresponding values.
        ///</summary>
        public IReadOnlyDictionary<DocumentPropertyName, string> DocumentProperties => CorePropertyHelpers.Get(this.Package);

        /// <summary>
        /// Set a known core property value
        /// </summary>
        /// <param name="name">Property to set</param>
        /// <param name="value">Value to assign</param>
        public void SetPropertyValue(DocumentPropertyName name, string value)
        {
            ThrowIfObjectDisposed();
            _ = CorePropertyHelpers.SetValue(Package, name.GetEnumName(), value);
            UpdateComplexFieldUsage(name.ToString().ToUpper(), value);
        }

        /// <summary>
        /// Returns a list of custom properties in this document.
        /// </summary>
        public IReadOnlyDictionary<string, object> CustomProperties
        {
            get
            {
                ThrowIfObjectDisposed();
                return CustomPropertyHelpers.Get(Package);
            }
        }

        /// <summary>
        /// Add a custom property to this document.
        /// If a custom property already exists with the same name it will be replace.
        /// CustomProperty names are case insensitive.
        /// </summary>
        public void AddCustomProperty(string name, string value) => AddCustomProperty(name, CustomProperty.LPWSTR, value);

        /// <summary>
        /// Add a custom property to this document.
        /// If a custom property already exists with the same name it will be replace.
        /// CustomProperty names are case insensitive.
        /// </summary>
        public void AddCustomProperty(string name, double value) => AddCustomProperty(name, CustomProperty.R8, value);

        /// <summary>
        /// Add a custom property to this document.
        /// If a custom property already exists with the same name it will be replace.
        /// CustomProperty names are case insensitive.
        /// </summary>
        public void AddCustomProperty(string name, bool value) => AddCustomProperty(name, CustomProperty.BOOL, value);

        /// <summary>
        /// Add a custom property to this document.
        /// If a custom property already exists with the same name it will be replace.
        /// CustomProperty names are case insensitive.
        /// </summary>
        public void AddCustomProperty(string name, DateTime value) => AddCustomProperty(name, CustomProperty.FILETIME, value.ToUniversalTime());

        /// <summary>
        /// Add a custom property to this document.
        /// If a custom property already exists with the same name it will be replace.
        /// CustomProperty names are case insensitive.
        /// </summary>
        public void AddCustomProperty(string name, int value) => AddCustomProperty(name, CustomProperty.I4, value);

        /// <summary>
        /// Add a custom property to this document.
        /// If a custom property already exists with the same name it will be replace.
        /// CustomProperty names are case insensitive.
        /// </summary>
        private void AddCustomProperty(string name, string type, object value)
        {
            ThrowIfObjectDisposed();
            if (CustomPropertyHelpers.Add(Package, name, type, value))
            {
                UpdateComplexFieldUsage(name.ToUpperInvariant(), value?.ToString()??"");
            }
        }

        /// <summary>
        /// Add an Image into this document from a fully qualified or relative filename.
        /// </summary>
        /// <param name="imageFileName">The fully qualified or relative filename.</param>
        /// <returns>An Image file.</returns>
        public Image AddImage(string imageFileName)
        {
            // The extension this file has will be taken to be its format.
            string contentType = Path.GetExtension(imageFileName) switch
            {
                ".tiff" => "image/tif",
                ".tif" => "image/tif",
                ".png" => "image/png",
                ".bmp" => "image/png",
                ".gif" => "image/gif",
                ".jpg" => "image/jpg",
                ".jpeg" => "image/jpeg",
                ".svg" => "image/svg",
                _ => "image/jpg",
            };

            using var fs = new FileStream(imageFileName, FileMode.Open, FileAccess.Read);
            return AddImage(fs, contentType);
        }

        /// <summary>
        /// Add an Image into this document from a Stream.
        /// </summary>
        /// <param name="imageStream">A Stream stream.</param>
        /// <param name="contentType">Content type - image/jpg</param>
        /// <returns>An Image file.</returns>
        public Image AddImage(Stream imageStream, string contentType = "image/jpg")
        {
            ThrowIfObjectDisposed();

            if (imageStream == null)
                throw new ArgumentNullException(nameof(imageStream));

            // Get all image parts in word\document.xml
            var relationshipsByImages = PackagePart.GetRelationshipsByType($"{Namespace.RelatedDoc.NamespaceName}/image");
            var imageParts = relationshipsByImages.Select(ir => Package.GetParts()
                        .FirstOrDefault(p => p.Uri.ToString().EndsWith(ir.TargetUri.ToString())))
                        .Where(e => e != null)
                        .ToList();

            foreach (var relsPart in Package.GetParts().Where(part => part.Uri.ToString().Contains("/word/")
                                            && part.ContentType.Equals(DocxContentType.Relationships)))
            {
                var relsPartContent = relsPart.Load();
                if (relsPartContent.Root == null)
                    throw new InvalidOperationException("Relationships missing root element.");
                var imageRelationships = relsPartContent.Root.Elements()
                                            .Where(imageRel => imageRel.Attribute("Type")
                                                    .Value.Equals($"{Namespace.RelatedDoc.NamespaceName}/image"));

                foreach (var imageRelationship in imageRelationships.Where(e => e.Attribute("Target") != null))
                {
                    string targetMode = imageRelationship.AttributeValue("TargetMode");
                    if (!targetMode.Equals("External"))
                    {
                        string imagePartUri = Path.Combine(Path.GetDirectoryName(relsPart.Uri.ToString()), imageRelationship.AttributeValue("Target"));
                        imagePartUri = Path.GetFullPath(imagePartUri.Replace("\\_rels", string.Empty));
                        imagePartUri = imagePartUri.Replace(Path.GetFullPath("\\"), string.Empty).Replace("\\", "/");

                        if (!imagePartUri.StartsWith("/"))
                        {
                            imagePartUri = "/" + imagePartUri;
                        }

                        imageParts.Add(Package.GetPart(new Uri(imagePartUri, UriKind.Relative)));
                    }
                }
            }

            // Loop through each image part in this document.
            foreach (var pp in imageParts)
            {
                using var existingImageStream = pp.GetStream(FileMode.Open, FileAccess.Read);

                // Compare this image to the new image being added. If it's the same file,
                // then reuse the existing Image resource rather than adding it again.
                if (HelperFunctions.IsSameFile(existingImageStream, imageStream))
                {
                    // Get the image object for this image part
                    string id = PackagePart.GetRelationshipsByType($"{Namespace.RelatedDoc.NamespaceName}/image")
                        .Where(r => r.TargetUri == pp.Uri)
                        .Select(r => r.Id).First();

                    // Return the Image object
                    return Images.First(i => i.Id == id);
                }
            }

            string imgPartUriPath;
            string extension = contentType.Substring(contentType.LastIndexOf("/", StringComparison.Ordinal) + 1);

            // Get a unique imgPartUriPath
            do
            {
                imgPartUriPath = $"/word/media/{Guid.NewGuid()}.{extension}";
            } while (Package.PartExists(new Uri(imgPartUriPath, UriKind.Relative)));

            var imagePackagePart = Package.CreatePart(new Uri(imgPartUriPath, UriKind.Relative), contentType, CompressionOption.Normal);

            // Create a new image relationship
            var imageRelation = PackagePart.CreateRelationship(imagePackagePart.Uri, TargetMode.Internal, $"{Namespace.RelatedDoc.NamespaceName}/image");

            // Open a Stream to the newly created Image part.
            using (var imageWriter = imagePackagePart.GetStream(FileMode.Create, FileAccess.Write))
            {
                // Using the Stream to the real image, copy this streams data into the newly create Image part.
                using (imageStream)
                {
                    if (imageStream.CanSeek)
                    {
                        imageStream.Seek(0, SeekOrigin.Begin);
                        imageStream.CopyTo(imageWriter);
                    }
                    else
                    {
                        byte[] bytes = new byte[imageStream.Length];
                        imageStream.Read(bytes, 0, (int)imageStream.Length);
                        imageWriter.Write(bytes, 0, (int)imageStream.Length);
                    }
                }
            }

            return new Image(this, imageRelation);
        }

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateFilePath">The path to the document template file.</param>
        ///<exception cref="FileNotFoundException">The document template file not found.</exception>
        public void ApplyTemplate(string templateFilePath) => ApplyTemplate(templateFilePath, true);

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateFilePath">The path to the document template file.</param>
        ///<param name="includeContent">Whether to copy the document template text content to document.</param>
        ///<exception cref="FileNotFoundException">The document template file not found.</exception>
        public void ApplyTemplate(string templateFilePath, bool includeContent)
        {
            if (!File.Exists(templateFilePath))
                throw new FileNotFoundException($"File could not be found {templateFilePath}", templateFilePath);

            using var packageStream = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read);
            ApplyTemplate(packageStream, includeContent);
        }

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateStream">The stream of the document template file.</param>
        public void ApplyTemplate(Stream templateStream) => ApplyTemplate(templateStream, true);

        ///<summary>
        /// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
        ///</summary>
        ///<param name="templateStream">The stream of the document template file.</param>
        ///<param name="includeContent">Whether to copy the document template text content to document.</param>
        public void ApplyTemplate(Stream templateStream, bool includeContent)
        {
            ThrowIfObjectDisposed();

            var templatePackage = Package.Open(templateStream);
            try
            {
                PackagePart templateDocumentPart = null;
                XDocument templateDocument = null;

                foreach (var packagePart in templatePackage.GetParts())
                {
                    switch (packagePart.Uri.ToString())
                    {
                        case "/word/document.xml":
                            templateDocumentPart = packagePart;
                            templateDocument = packagePart.Load();
                            break;

                        case "/_rels/.rels":
                            if (!Package.PartExists(packagePart.Uri))
                                Package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption);

                            var globalRelationshipsPart = Package.GetPart(packagePart.Uri);
                            using (var tr = new StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8))
                            {
                                using var tw = new StreamWriter(globalRelationshipsPart.GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8);
                                tw.Write(tr.ReadToEnd());
                            }
                            break;

                        case "/word/_rels/document.xml.rels":
                            break;

                        default:
                            if (!Package.PartExists(packagePart.Uri))
                                Package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption);
                            var encoding = packagePart.Uri.ToString().EndsWith(".xml") ||
                                           packagePart.Uri.ToString().EndsWith(".rels") ? Encoding.UTF8 : Encoding.Default;

                            var nativePart = Package.GetPart(packagePart.Uri);
                            using (var tr = new StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read), encoding))
                            {
                                using var tw = new StreamWriter(nativePart.GetStream(FileMode.Create, FileAccess.Write), tr.CurrentEncoding);
                                tw.Write(tr.ReadToEnd());
                            }
                            break;
                    }
                }

                if (templateDocumentPart != null)
                {
                    string mainContentType = templateDocumentPart.ContentType.Replace("template.main", "document.main");
                    if (Package.PartExists(templateDocumentPart.Uri))
                        Package.DeletePart(templateDocumentPart.Uri);

                    var documentNewPart = Package.CreatePart(templateDocumentPart.Uri, mainContentType, templateDocumentPart.CompressionOption);
                    using (var writer = XmlWriter.Create(documentNewPart.GetStream(FileMode.Create, FileAccess.Write)))
                    {
                        templateDocument.WriteTo(writer);
                    }

                    foreach (var documentPartRel in templateDocumentPart.GetRelationships())
                    {
                        documentNewPart.CreateRelationship(documentPartRel.TargetUri, documentPartRel.TargetMode, documentPartRel.RelationshipType, documentPartRel.Id);
                    }

                    // TODO:
                    //PackagePart = documentNewPart;
                    //mainDoc = templateDocument;
                    OnLoadDocument();
                }

                if (!includeContent)
                {
                    foreach (var paragraph in Paragraphs)
                    {
                        paragraph.Remove();
                    }
                }
            }
            finally
            {
                Package.Flush();
                templatePackage.Close();
            }
        }

        /// <summary>
        /// Close and release all resources used by this document.
        /// </summary>
        public void Close()
        {
            Package?.Close();
            memoryStream?.Close();
            Package = null;
            memoryStream = null;

            filename = null;
            stream = null;
        }

        /// <summary>
        /// Releases all resources used by this document.
        /// </summary>
        void IDisposable.Dispose()
        {
            Close();
        }

        /// <summary>
        /// Insert a chart in document
        /// </summary>
        public void InsertChart(Chart chart)
        {
            ThrowIfObjectDisposed();

            // Create a new chart part uri.
            string chartPartUriPath;
            int chartIndex = 0;

            do
            {
                chartIndex++;
                chartPartUriPath = $"/word/charts/chart{chartIndex}.xml";
            } while (Package.PartExists(new Uri(chartPartUriPath, UriKind.Relative)));

            // Create chart part.
            var chartPackagePart = Package.CreatePart(new Uri(chartPartUriPath, UriKind.Relative), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", CompressionOption.Normal);

            // Create a new chart relationship
            string id = GetNextRelationshipId();
            _ = PackagePart.CreateRelationship(chartPackagePart.Uri, TargetMode.Internal, $"{Namespace.RelatedDoc.NamespaceName}/chart", id);

            // Save a chart info the chartPackagePart
            chartPackagePart.Save(chart.Xml);

            // Insert a new chart into a paragraph.
            var p = this.AddParagraph();
            var chartElement = new XElement(Name.Run,
                new XElement(Namespace.Main + "drawing",
                    new XElement(Namespace.WordProcessingDrawing + "inline",
                        new XElement(Namespace.WordProcessingDrawing + "extent",
                            new XAttribute("cx", 5486400),
                            new XAttribute("cy", 3200400)),
                        new XElement(Namespace.WordProcessingDrawing + "effectExtent",
                            new XAttribute("l", 0),
                            new XAttribute("t", 0),
                            new XAttribute("r", 19050),
                            new XAttribute("b", 19050)),
                        new XElement(Namespace.WordProcessingDrawing + "docPr",
                            new XAttribute("id", 1),
                            new XAttribute("name", "chart")),
                        new XElement(Namespace.DrawingMain + "graphic",
                            new XElement(Namespace.DrawingMain + "graphicData",
                                new XAttribute("uri", Namespace.Chart.NamespaceName),
                                new XElement(Namespace.Chart + "chart",
                                    new XAttribute(Namespace.RelatedDoc + "id", id)
                                )
                            )
                        )
                    )
               ));
            p.Xml.Add(chartElement);
        }

        /// <summary>
        /// Inserts a default TOC into the current document.
        /// Title: Table of contents
        /// Switches will be: TOC \h \o '1-3' \u \z
        /// </summary>
        /// <returns>The inserted TableOfContents</returns>
        public TableOfContents InsertDefaultTableOfContents()
        {
            return InsertTableOfContents("Table of contents",
                TableOfContentsSwitches.O
              | TableOfContentsSwitches.H
              | TableOfContentsSwitches.Z
              | TableOfContentsSwitches.U,
                HeadingType.Heading1.GetEnumName());
        }

        /// <summary>
        /// Inserts a table of contents into the document.
        /// </summary>
        /// <param name="title">The title of the TOC</param>
        /// <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
        /// <param name="headerStyle">Lets you set the style name of the TOC header</param>
        /// <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
        /// <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
        /// <returns>The inserted TableOfContents</returns>
        public TableOfContents InsertTableOfContents(string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null)
        {
            ThrowIfObjectDisposed();

            var toc = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);

            XElement sectPr = Xml.Elements(Name.SectionProperties).SingleOrDefault();
            if (sectPr != null)
            {
                sectPr.AddBeforeSelf(toc.Xml);
            }
            else
            {
                Xml.Add(toc.Xml);
            }

            return toc;
        }

        /// <summary>
        /// Inserts at TOC into the current document before the provided paragraph.
        /// </summary>
        /// <param name="beforeMe">The paragraph to insert the ToC before</param>
        /// <param name="title">The title of the TOC</param>
        /// <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
        /// <param name="headerStyle">Lets you set the style name of the TOC header</param>
        /// <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
        /// <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
        /// <returns>The inserted TableOfContents</returns>
        public TableOfContents InsertTableOfContents(Paragraph beforeMe, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null)
        {
            ThrowIfObjectDisposed();

            if (beforeMe == null)
                throw new ArgumentNullException(nameof(beforeMe));

            var toc = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);
            beforeMe.Xml.AddBeforeSelf(toc.Xml);

            return toc;
        }

        /// <summary>
        /// Save this document back to the location it was loaded from.
        /// </summary>
        public void Save()
        {
            ThrowIfObjectDisposed();

            if (filename == null && stream == null)
            {
                throw new InvalidOperationException(
                    "No filename or stream to save to - use SaveAs to persist document.");
            }

            // Save the main document
            PackagePart.Save(mainDoc);

            // Refresh settings
            settingsDoc = settingsPart.Load();

            // Bump the revision and add it to the settings document.
            settingsDoc.Root!.Element(Namespace.Main + "rsids")
                            ?.Add(new XElement(Namespace.Main + "rsid",
                                 new XAttribute(Name.MainVal, HelperFunctions.GenerateRevisionStamp(RevisionId, out revision))));

            // Save all the sections
            Sections.ToList().ForEach(s =>
            {
                s.Headers.Save();
                s.Footers.Save();
            });

            Styles?.Save();
            NumberingStyles?.Save();

            settingsPart.Save(settingsDoc);
            endnotesPart?.Save(endnotesDoc);
            footnotesPart?.Save(footnotesDoc);

            // Close the package and commit changes to the memory stream.
            // Note that .NET core requires we close the package - not just flush it.
            Package.Close();

            // Save back to the file or stream
            if (filename != null)
            {
                memoryStream.Seek(0, SeekOrigin.Begin);
                using var fs = File.Create(filename);
                memoryStream.WriteTo(fs);
            }
            else //if (stream != null)
            {
                if (stream.CanSeek)
                {
                    stream.SetLength(0);
                    stream.Seek(0, SeekOrigin.Begin);
                }
                memoryStream.WriteTo(stream);
            }

            // Reopen the package.
            Package = Package.Open(memoryStream, FileMode.Open, FileAccess.ReadWrite);
            RefreshDocumentParts();
        }

        /// <summary>
        /// Save this document to a file.
        /// </summary>
        /// <param name="newFileName">The filename to save this document as.</param>
        public void SaveAs(string newFileName)
        {
            stream = null;
            filename = newFileName;
            Save();
        }

        /// <summary>
        /// Save this document to a Stream.
        /// </summary>
        /// <param name="newStreamDestination">The Stream to save this document to.</param>
        public void SaveAs(Stream newStreamDestination)
        {
            filename = null;
            stream = newStreamDestination;
            Save();
        }

        /// <summary>
        /// Load the properties from the loaded package
        /// </summary>
        /// <param name="package">Loaded package</param>
        /// <returns>New Docx object</returns>
        internal static DocX CreateDocumentFromPackage(Package package)
        {
            var document = new DocX { Package = package };
            document.Document = document;
            document.OnLoadDocument();

            return document;
        }

        /// <summary>
        /// Adds a new styles.xml to the package and relates it to this document.
        /// </summary>
        internal void AddDefaultStyles()
        {
            // If the document contains no /word/styles.xml create one and associate it
            if (!Package.PartExists(Relations.Styles.Uri))
            {
                var stylesDoc = HelperFunctions.AddDefaultStylesXml(Package, out var stylesPart);
                if (stylesDoc.Element(Namespace.Main + "styles") == null)
                    throw new Exception("Missing root styles collection.");

                Styles = new StyleManager(this, stylesPart);
            }

            Debug.Assert(Styles != null);
        }

        /// <summary>
        /// Adds the hyperlink style to the document. This is done the first time
        /// a hyperlink is added. If the style already exists, this method does nothing.
        /// </summary>
        internal void AddHyperlinkStyle()
        {
            ThrowIfObjectDisposed();
            AddDefaultStyles();

            if (!Styles.HasStyle("Hyperlink", StyleType.Character))
                Styles.Add(Resource.HyperlinkStyle(RevisionId));
        }

        /// <summary>
        /// Recreate the links to the different package parts when we're re-creating the package.
        /// </summary>
        internal void RefreshDocumentParts()
        {
            ThrowIfObjectDisposed();

            // Get the main document part
            PackagePart = Package.GetParts().Single(p =>
                p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase) ||
                p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

            // Load the settings
            settingsPart = Package.GetPart(Relations.Settings.Uri);

            // Load all the sections
            foreach (var rel in PackagePart.GetRelationships())
            {
                if (rel.RelationshipType == Relations.Endnotes.RelType)
                    endnotesPart = Package.GetPart(Relations.Endnotes.Uri);
                else if (rel.RelationshipType == Relations.Footnotes.RelType)
                    footnotesPart = Package.GetPart(Relations.Footnotes.Uri);
                else if (rel.RelationshipType == Relations.Styles.RelType)
                    Styles = new StyleManager(this, Package.GetPart(Relations.Styles.Uri));
                else if (rel.RelationshipType == Relations.Numbering.RelType)
                    NumberingStyles = new NumberingStyleManager(this, Package.GetPart(Relations.Numbering.Uri));
            }
        }

        /// <summary>
        /// Loads the document contents from the assigned package
        /// </summary>
        internal void OnLoadDocument()
        {
            if (Package == null)
                throw new InvalidOperationException($"Cannot load package parts when {nameof(Package)} property is not set.");

            // Load all the package parts
            RefreshDocumentParts();

            // Grab the main document
            mainDoc = PackagePart.Load();

            // Set the DocElement XML value
            Xml = mainDoc.Root!.Element(Namespace.Main + "body");
            if (Xml == null)
                throw new InvalidOperationException($"Missing {Namespace.Main + "body"} expected content.");

            // Get the last revision id
            string revValue = Xml.GetSectionProperties().RevisionId;
            revision = uint.Parse(revValue, System.Globalization.NumberStyles.AllowHexSpecifier);
            revision++; // bump revision

            // Load all the XML files
            settingsDoc = settingsPart?.Load();
            endnotesDoc = endnotesPart?.Load();
            footnotesDoc = footnotesPart?.Load();
        }

        /// <summary>
        /// This adds the default numbering.xml to the document if it doesn't exist.
        /// </summary>
        internal void AddDefaultNumberingPart()
        {
            // If we don't have any numbering styles in this document, add a default document.
            if (NumberingStyles == null)
            {
                var packagePart = Package.CreatePart(Relations.Numbering.Uri, Relations.Numbering.ContentType, CompressionOption.Maximum);
                var template = Resource.NumberingXml();
                packagePart.Save(template);
                PackagePart.CreateRelationship(packagePart.Uri, TargetMode.Internal, Relations.Numbering.RelType);
                NumberingStyles = new NumberingStyleManager(this, packagePart);
            }

            // See if we have the list style
            AddDefaultStyles();
            if (Styles.HasStyle("ListParagraph", StyleType.Paragraph))
            {
                Styles.Add(Resource.ListParagraphStyle(RevisionId));
            }
        }

        /// <summary>
        /// Create a new relationship id by locating the last one used.
        /// </summary>
        /// <returns></returns>
        private string GetNextRelationshipId()
        {
            // Last used id (0 if none)
            int id = PackagePart.GetRelationships()
                        .Where(r => r.Id.Substring(0, 3).Equals("rId"))
                        .Select(r => int.TryParse(r.Id.Substring(3), out int result) ? result : 0)
                        .DefaultIfEmpty()
                        .Max();
            return $"rId{id + 1}";
        }

        /// <summary>
        /// Method to throw an ObjectDisposedException
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void ThrowIfObjectDisposed()
        {
            if (Package == null)
                throw new ObjectDisposedException("DocX object has been disposed.");
        }

        /// <summary>
        /// Renumber the tracking ids
        /// </summary>
        internal void RenumberIds()
        {
            ThrowIfObjectDisposed();

            var trackerIds = mainDoc.Descendants()
                    .Where(d => d.Name.LocalName == "ins" || d.Name.LocalName == "del")
                    .Select(d => d.Attribute(Name.Id))
                    .ToList();

            for (int i = 0; i < trackerIds.Count; i++)
            {
                trackerIds[i].Value = i.ToString();
            }
        }

        /// <summary>
        /// Update all usages of a complex field.
        /// </summary>
        /// <param name="name">Name of the core property</param>
        /// <param name="value">Value</param>
        private void UpdateComplexFieldUsage(string name, string value)
        {
            ThrowIfObjectDisposed();

            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));

            string matchPattern = $@"{name.ToUpperInvariant()} \\\* MERGEFORMAT";

            // Update in all sections.
            var documents = new[] {mainDoc.Root}
                .Union(Sections.SelectMany(s => s.Headers).Select(h => h.Xml))
                .Union(Sections.SelectMany(s => s.Footers).Select(f => f.Xml));

            foreach (var doc in documents)
            {
                // Look for all the instrText fields matching this name.
                foreach (var e in doc.Descendants(Namespace.Main + "instrText")
                            .Where(e => Regex.IsMatch(e.Value.ToUpperInvariant(), matchPattern)))
                {
                    // Back up to <w:r> parent
                    var node = e.Parent;
                    if (node?.Name != Name.Run)
                        continue;

                    var paragraph = node.Parent;
                    if (paragraph == null)
                        continue;
                    
                    // Walk down to the next w:r/w:t node.
                    while (node?.Parent == paragraph)
                    {
                        if (node.Element(Name.Text) != null)
                            break;
                        node = node.NextNode as XElement;
                    }

                    // Didn't find it.
                    if (node?.Parent != paragraph)
                        continue;

                    // Replace the text.
                    node.SetElementValue(Name.Text, value);
                }
            }
        }

        /// <summary>
        /// This method marks the placeholder fields as invalid and ensures Word updates them
        /// when it next loads the document.
        /// </summary>
        internal void InvalidatePlaceholderFields()
        {
            if (!settingsDoc.Descendants(Namespace.Main + "updateFields").Any())
            {
                settingsDoc.Root!.Add(new XElement(Namespace.Main + "updateFields",
                                           new XAttribute(Name.MainVal, true)));
            }
        }

        /// <summary>
        /// Get a document ID to use with inserted XML. This should always be unique within the document.
        /// </summary>
        /// <returns>New document id</returns>
        internal long GetNextDocumentId()
        {
            if (nextDocumentId == 0)
                nextDocumentId = HelperFunctions.FindLastUsedDocId(mainDoc);
            return ++nextDocumentId;
        }

        /// <summary>
        /// Create a new Picture.
        /// </summary>
        /// <param name="rid">A unique id that identifies an Image embedded in this document.</param>
        /// <param name="name">The name of this Picture.</param>
        /// <param name="description">The description of this Picture.</param>
        internal Picture CreatePicture(string rid, string name, string description)
        {
            if (string.IsNullOrWhiteSpace(rid))
            {
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(rid));
            }

            long id = GetNextDocumentId();
            int cx, cy;

            var relationship = PackagePart.GetRelationship(rid);
            using (var partStream = Package.GetPart(relationship.TargetUri).GetStream())
            using (var img = System.Drawing.Image.FromStream(partStream))
            {
                cx = img.Width * 9526;
                cy = img.Height * 9526;
            }

            return new Picture(this, Resource.DrawingElement(id, name, description, cx, cy, rid),
                                new Image(this, relationship));
        }

        /// <summary>
        /// Ensure a _rels XML file has been created for the given package part.
        /// Most documents just have document.rels, but more complex documents can
        /// include other relationship files.
        /// </summary>
        /// <param name="part">Part to lookup</param>
        /// <returns>URI for the given part</returns>
        internal Uri EnsureRelsPathExists(PackagePart part)
        {
            // Convert the path of this mainPart to its equivalent rels file path.
            string path = part.Uri.OriginalString.Replace("/word/", "");
            Uri relationshipPath = new Uri($"/word/_rels/{path}.rels", UriKind.Relative);

            // Check to see if the rels file exists and create it if not.
            if (!Document.Package.PartExists(relationshipPath))
            {
                PackagePart pp = Document.Package.CreatePart(relationshipPath, DocxContentType.Relationships, CompressionOption.Maximum);
                pp.Save(new XDocument(
                    new XDeclaration("1.0", "UTF-8", "yes"),
                    new XElement(Namespace.RelatedPackage + "Relationships")
                ));
            }

            return relationshipPath;
        }

        /// <summary>
        /// Loads a header/footer document from the related package.
        /// </summary>
        /// <param name="id">Relationship id</param>
        /// <param name="part">Returned packagePart document is loaded from</param>
        /// <param name="doc">Document</param>
        internal void FindHeaderFooterById(string id, out PackagePart part, out XDocument doc)
        {
            // Get the Xml file for this Header or Footer. Each one is saved into a different
            // document and kept as a relationship in the document.
            Uri partUri = PackagePart.GetRelationship(id).TargetUri;
            if (!partUri.OriginalString.StartsWith("/word/"))
            {
                partUri = new Uri("/word/" + partUri.OriginalString, UriKind.Relative);
            }

            // Get the PackagePart and return the XM document.
            part = Package.GetPart(partUri);

            // We keep the header/footer XML documents in memory so we don' thave
            // to save them back to the package every time a change is made.
            if (!headerFooterCache.TryGetValue(id, out doc))
            {
                doc = part.Load();
                AdjustHeaderFooterCache(id, doc);
            }
        }

        /// <summary>
        /// Add or remove a header/footer to/from the cache
        /// </summary>
        /// <param name="id">Relationship id in the main document</param>
        /// <param name="doc">XML document for the header/footer</param>
        internal void AdjustHeaderFooterCache(string id, XDocument doc)
        {
            if (doc != null)
                headerFooterCache.Add(id, doc);
            else
                headerFooterCache.Remove(id);
        }
        
        /// <summary>
        /// Locate a paragraph from a character index.
        /// </summary>
        /// <param name="index"></param>
        /// <returns>Paragraph</returns>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        internal Paragraph FindParagraphByIndex(int index)
        {
            if (index < 0)
                throw new ArgumentOutOfRangeException(nameof(index));

            // If the insertion position is first (0) and there are no paragraphs, then return null.
            var lookup = Paragraphs.ToDictionary(paragraph => paragraph.EndIndex);
            if (lookup.Keys.Count == 0 && index == 0)
            {
                return null;
            }

            // Find the paragraph that contains the index
            foreach (int paragraphEndIndex in lookup.Keys.Where(paragraphEndIndex => paragraphEndIndex >= index))
            {
                return lookup[paragraphEndIndex];
            }

            throw new ArgumentOutOfRangeException(nameof(index));
        }
    }
}