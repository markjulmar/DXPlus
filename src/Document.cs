using System.Globalization;
using DXPlus.Charts;
using DXPlus.Helpers;
using DXPlus.Resources;
using System.IO.Packaging;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DXPlus.Internal;

[assembly:InternalsVisibleTo("DXPlus.Tests")]

namespace DXPlus;

/// <summary>
/// Represents a document.
/// </summary>
public sealed class Document : BlockContainer, IDocument
{
    /// <summary>
    /// Default constructor (private)
    /// </summary>
    private Document()
    {
    }

    /// <summary>
    /// Loads a Word document from a Stream.
    /// </summary>
    /// <param name="stream">The Stream to load the document from.</param>
    /// <returns>
    /// Returns an IDocument object which represents the document.
    /// </returns>
    public static IDocument Load(Stream stream)
    {
        var ms = new MemoryStream();
        stream.Seek(0, SeekOrigin.Begin);
        stream.CopyTo(ms);

        var document = new Document();

        document.documentPackage = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);
        document.memoryStream = ms;
        document.stream = stream;

        document.LoadDocumentParts();

        return document;
    }

    /// <summary>
    /// Loads a Word document from a fully qualified or relative filename.
    /// </summary>
    /// <param name="filename">The fully qualified or relative filename.</param>
    /// <returns>
    /// Returns an IDocument object which represents the document.
    /// </returns>
    public static IDocument Load(string filename)
    {
        if (string.IsNullOrWhiteSpace(filename))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(filename));
        if (!File.Exists(filename))
            throw new FileNotFoundException("Docx file doesn't exist", filename);

        // Load the file from disk into memory.
        // Note: keep memory stream OPEN as it represents our document and the
        // underlying PackageManager will use this stream to load various parts
        // as we access the document.
        var ms = new MemoryStream();
        using (var fs = new FileStream(filename, FileMode.Open))
        {
            fs.CopyTo(ms);
        }

        // Create the document
        var document = new Document();

        document.documentPackage = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);
        document.filename = filename;
        document.memoryStream = ms;

        document.LoadDocumentParts();

        return document;
    }

    /// <summary>
    /// Create a new Word document
    /// </summary>
    /// <param name="filename">Optional filename</param>
    /// <returns>New document</returns>
    public static IDocument Create(string? filename = null) => Create(filename, DocumentTypes.Document);

    /// <summary>
    /// Create a new Word document template
    /// </summary>
    /// <param name="filename">Optional filename</param>
    /// <returns>New template</returns>
    public static IDocument CreateTemplate(string? filename = null) => Create(filename, DocumentTypes.Template);

    /// <summary>
    /// Method to create an unnamed document
    /// </summary>
    /// <param name="filename">optional filename</param>
    /// <param name="documentType">Type to create</param>
    /// <returns>New document</returns>
    private static Document Create(string? filename, DocumentTypes documentType)
    {
        var doc = (Document)Load(HelperFunctions.CreateDocumentType(documentType));
        doc.stream!.Dispose();
        doc.stream = null;
        doc.filename = filename;
        doc.revision--;

        return doc;
    }

    /// <summary>
    /// Editing session id for this session.
    /// </summary>
    public string RevisionId => revision.ToString("X8");

    /// <summary>
    /// Retrieve all bookmarks
    /// </summary>
    public BookmarkCollection Bookmarks => new(Paragraphs.SelectMany(p => p.GetBookmarks()));

    /// <summary>
    /// Numbering styles in this document.
    /// </summary>
    public NumberingStyleManager NumberingStyles => numberingStyles ??= AddDefaultNumberingPart();

    /// <summary>
    /// The style manager for this document.
    /// </summary>
    public StyleManager Styles => styleManager ??= AddDefaultStyles();

    /// <summary>
    /// Returns all comments in the document
    /// </summary>
    public IEnumerable<Comment> Comments
        => commentManager == null ? Enumerable.Empty<Comment>() : commentManager.Comments;

    /// <summary>
    /// Retrieve a specific comment by id.
    /// </summary>
    /// <param name="id">Numeric id</param>
    /// <returns>Comment or null if no comment exists.</returns>
    public Comment? GetComment(int id) => Comments.SingleOrDefault(c => c.Id == id);

    /// <summary>
    /// Returns all comments for a specific author in the document.
    /// </summary>
    /// <param name="authorName">Author name</param>
    public IEnumerable<Comment> CommentsBy(string authorName) => Comments.Where(c =>
        string.Compare(authorName, c.AuthorName, StringComparison.OrdinalIgnoreCase) == 0);

    /// <summary>
    /// Returns all the reviewers for the document
    /// </summary>
    public IEnumerable<string> Reviewers => Comments.Select(c => c.AuthorName).Distinct();

    /// <summary>
    /// True if the Document use different Headers and Footers for odd and even pages.
    /// </summary>
    public bool DifferentEvenOddHeadersFooters
    {
        get => settingsDoc!.Root!.Element(Namespace.Main + "evenAndOddHeaders") != null;

        set
        {
            var evenAndOddHeaders = settingsDoc!.Root!.Element(Namespace.Main + "evenAndOddHeaders");
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
        => endnotesDoc?.Root?.Elements(Namespace.Main + "endnote").Select(HelperFunctions.GetText) ?? Enumerable.Empty<string>();

    /// <summary>
    /// Get the text of each footnote from this document
    /// </summary>
    public IEnumerable<string> FootnotesText
        => footnotesDoc?.Root?.Elements(Namespace.Main + "footnote").Select(HelperFunctions.GetText) ?? Enumerable.Empty<string>();

    /// <summary>
    /// Returns a list of Images in this document.
    /// </summary>
    public IReadOnlyList<Image> Images
    {
        get
        {
            ThrowIfNoPackage();
            var imageRelationships = PackagePart.GetRelationshipsByType($"{Namespace.RelatedDoc.NamespaceName}/image");
            return (imageRelationships.Any()
                ? imageRelationships.Select(i => new Image(this, i)).ToList()
                : new List<Image>()).AsReadOnly();
        }
    }

    ///<summary>
    /// Returns the list of document properties with corresponding values.
    ///</summary>
    public IReadOnlyDictionary<DocumentPropertyName, string> DocumentProperties
    {
        get
        {
            ThrowIfNoPackage();
            return CorePropertyHelpers.Get(Package);
        }
    }

    /// <summary>
    /// Set a known core property value
    /// </summary>
    /// <param name="name">Property to set</param>
    /// <param name="value">Value to assign</param>
    public void SetPropertyValue(DocumentPropertyName name, string value)
    {
        ThrowIfNoPackage();
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
            ThrowIfNoPackage();
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
        ThrowIfNoPackage();

        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
        if (string.IsNullOrWhiteSpace(type))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(type));
        if (value == null) throw new ArgumentNullException(nameof(value));

        if (CustomPropertyHelpers.Add(Package, name, type, value))
        {
            UpdateComplexFieldUsage(name.ToUpperInvariant(), value.ToString()??"");
        }
    }

    /// <summary>
    /// Add an Image into this document from a fully qualified or relative filename.
    /// </summary>
    /// <param name="imageFileName">The fully qualified or relative filename.</param>
    /// <param name="contentType">Content type</param>
    /// <returns>An Image file.</returns>
    public Image AddImage(string imageFileName, string? contentType)
    {
        ThrowIfNoPackage();

        if (imageFileName is null)
            throw new ArgumentNullException(nameof(imageFileName));

        if (!File.Exists(imageFileName))
            throw new ArgumentException("Missing image file.", nameof(imageFileName));

        if (string.IsNullOrEmpty(contentType))
        {
            // The extension this file has will be taken to be its format.
            contentType = Path.GetExtension(imageFileName).ToLower() switch
            {
                ".tiff" => ImageContentType.Tiff,
                ".tif" => ImageContentType.Tiff,
                ".png" => ImageContentType.Png,
                ".bmp" => ImageContentType.Png,
                ".gif" => ImageContentType.Gif,
                ".jpg" => ImageContentType.Jpg,
                ".jpeg" => ImageContentType.Jpeg,
                ".svg" => ImageContentType.Svg,
                _ => throw new ArgumentException("Unable to determine content type from filename.", nameof(imageFileName)),
            };
        }

        contentType = contentType.ToLower();
        ValidateContentType(contentType);

        using var fs = new FileStream(imageFileName, FileMode.Open, FileAccess.Read);
        return AddImage(fs, contentType, Path.GetFileName(imageFileName));
    }

    /// <summary>
    /// Add an Image into this document from a Stream.
    /// </summary>
    /// <param name="imageStream">A stream with the image.</param>
    /// <param name="contentType">Content type to add</param>
    /// <returns>An Image file.</returns>
    public Image AddImage(Stream imageStream, string contentType)
        => AddImage(imageStream, contentType, "image");

    /// <summary>
    /// Validates the passed image content type
    /// </summary>
    /// <param name="contentType"></param>
    private static void ValidateContentType(string contentType)
    {
        if (typeof(ImageContentType).GetProperties(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public)
                .FirstOrDefault(pi => pi.GetValue(null)?.ToString() == contentType) != null)
            return;

        throw new ArgumentException("Bad content type - use one of the constants from ImageContentType.", nameof(contentType));
    }

    /// <summary>
    /// Add an Image into this document from a Stream.
    /// </summary>
    /// <param name="imageStream">A Stream stream.</param>
    /// <param name="contentType">Content type to add</param>
    /// <param name="imageFileName">Filename (if any)</param>
    /// <returns>An Image file.</returns>
    private Image AddImage(Stream imageStream, string contentType, string imageFileName)
    {
        ThrowIfNoPackage();

        if (imageStream == null)
            throw new ArgumentNullException(nameof(imageStream));
        if (string.IsNullOrEmpty(contentType))
            throw new ArgumentException($"'{nameof(contentType)}' cannot be null or empty.", nameof(contentType));

        contentType = contentType.ToLower();
        ValidateContentType(contentType);

        // See if the image is already in the document. If so, we'll
        // reuse the image.
        var existingImage = LocateExistingImageResource(imageStream);
        if (existingImage != null)
            return existingImage;

        // This is a new image which needs to be added to the rels document.
        string extension = contentType[(contentType.LastIndexOf("/", StringComparison.Ordinal) + 1)..];
        imageFileName = string.IsNullOrEmpty(imageFileName) ? "image" : Path.GetFileNameWithoutExtension(imageFileName);

        // Get a unique imgPartUriPath - start with the existing
        // filename and then append numeric to get something unique.
        string imgPartUriPath = $"/word/media/{imageFileName}.{extension}";
        if (Package.PartExists(new Uri(imgPartUriPath, UriKind.Relative)))
        {
            int i = 1;
            do
            {
                imgPartUriPath = $"/word/media/{imageFileName}{i}.{extension}";
                i++;
            } while (Package.PartExists(new Uri(imgPartUriPath, UriKind.Relative)));
        }

        // Create the package part
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
                    var bytes = new byte[imageStream.Length];
                    _ = imageStream.Read(bytes, 0, (int)imageStream.Length);
                    imageWriter.Write(bytes, 0, (int)imageStream.Length);
                }
            }
        }

        return new Image(this, imageRelation);
    }

    /// <summary>
    /// Locates an existing image resource by comparing bitmap images
    /// contained in the resource parts
    /// </summary>
    /// <param name="imageStream">Image stream to locate</param>
    /// <returns>Image or null if it's not in the document</returns>
    private Image? LocateExistingImageResource(Stream imageStream)
    {
        // Get all image parts in word\document.xml
        var relationshipsByImages = PackagePart.GetRelationshipsByType($"{Namespace.RelatedDoc.NamespaceName}/image");
        var imageParts = relationshipsByImages.Select(ir => Package.GetParts()
                .FirstOrDefault(p => p.Uri.ToString().EndsWith(ir.TargetUri.ToString())))
            .Where(e => e != null)
            .ToList();

        foreach (var relsPart in Package.GetParts()
                     .Where(part => part.Uri.ToString().Contains("/word/")
                                    && part.ContentType.Equals(DocxContentType.Relationships)))
        {
            var relsPartContent = relsPart.Load();
            if (relsPartContent.Root == null)
                throw new InvalidOperationException("Relationships missing root element.");
            var imageRelationships = relsPartContent.Root.Elements()
                .Where(imageRel => imageRel.Attribute("Type")?.Value.Equals($"{Namespace.RelatedDoc.NamespaceName}/image") == true);

            foreach (var imageRelationship in imageRelationships.Where(e => e.Attribute("Target") != null))
            {
                string? targetMode = imageRelationship.AttributeValue("TargetMode")?.ToLowerInvariant();
                if (targetMode != "external")
                {
                    string? prefix = Path.GetDirectoryName(relsPart.Uri.ToString());
                    string? target = imageRelationship.AttributeValue("Target");

                    if (target == null) continue;

                    string imagePartUri = prefix != null ? Path.Combine(prefix, target) : target;
                    imagePartUri = Path.GetFullPath(imagePartUri.Replace(@"\_rels", string.Empty));
                    imagePartUri = imagePartUri.Replace(Path.GetFullPath(@"\"), string.Empty).Replace('\\','/');

                    if (!imagePartUri.StartsWith('/'))
                    {
                        imagePartUri = $"/{imagePartUri}";
                    }

                    imageParts.Add(Package.GetPart(new Uri(imagePartUri, UriKind.Relative)));
                }
            }
        }

        // Loop through each image part in this document.
        foreach (var packagePart in imageParts.Where(part => part != null))
        {
            using var existingImageStream = packagePart!.GetStream(FileMode.Open, FileAccess.Read);

            // Compare this image to the new image being added. If it's the same file,
            // then reuse the existing Image resource rather than adding it again.
            if (HelperFunctions.IsSameFile(existingImageStream, imageStream))
            {
                // Get the image object for this image part
                string id = PackagePart.GetRelationshipsByType($"{Namespace.RelatedDoc.NamespaceName}/image")
                    .Where(r => r.TargetUri == packagePart.Uri)
                    .Select(r => r.Id).First();

                // Return the Image object
                return Images.First(i => i.Id == id);
            }
        }

        // No matching image
        return null;
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
        ThrowIfNoPackage();

        var templatePackage = Package.Open(templateStream);
        try
        {
            PackagePart? templateDocumentPart = null;
            XDocument? templateDocument = null;

            foreach (var part in templatePackage.GetParts())
            {
                switch (part.Uri.ToString())
                {
                    case "/word/document.xml":
                        templateDocumentPart = part;
                        templateDocument = part.Load();
                        break;

                    case "/_rels/.rels":
                        if (!Package.PartExists(part.Uri))
                            Package.CreatePart(part.Uri, part.ContentType, part.CompressionOption);

                        var globalRelationshipsPart = Package.GetPart(part.Uri);
                        using (var tr = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8))
                        {
                            using var tw = new StreamWriter(globalRelationshipsPart.GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8);
                            tw.Write(tr.ReadToEnd());
                        }
                        break;

                    case "/word/_rels/document.xml.rels":
                        break;

                    default:
                        if (!Package.PartExists(part.Uri))
                            Package.CreatePart(part.Uri, part.ContentType, part.CompressionOption);
                        var encoding = part.Uri.ToString().EndsWith(".xml") ||
                                       part.Uri.ToString().EndsWith(".rels") ? Encoding.UTF8 : Encoding.Default;

                        var nativePart = Package.GetPart(part.Uri);
                        using (var tr = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read), encoding))
                        {
                            using var tw = new StreamWriter(nativePart.GetStream(FileMode.Create, FileAccess.Write), tr.CurrentEncoding);
                            tw.Write(tr.ReadToEnd());
                        }
                        break;
                }
            }

            if (templateDocumentPart != null && templateDocument != null)
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

                LoadDocumentParts();
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
        Package.Close();
        memoryStream?.Close();

        documentPackage = null;
        memoryStream = null;
        filename = null;
        stream = null;
    }

    /// <summary>
    /// Releases all resources used by this document.
    /// </summary>
    void IDisposable.Dispose() => Close();

    /// <summary>
    /// Insert a chart in document
    /// </summary>
    public void InsertChart(Chart chart)
    {
        ThrowIfNoPackage();

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
    public TableOfContents InsertTableOfContents(string title, TableOfContentsSwitches switches, 
                    string? headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null)
    {
        ThrowIfNoPackage();

        var toc = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);

        var sectPr = Xml.Elements(Name.SectionProperties).SingleOrDefault();
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
    /// <param name="insertBefore">The paragraph to insert the ToC before</param>
    /// <param name="title">The title of the TOC</param>
    /// <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
    /// <param name="headerStyle">Lets you set the style name of the TOC header</param>
    /// <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
    /// <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
    /// <returns>The inserted TableOfContents</returns>
    public TableOfContents InsertTableOfContents(Paragraph insertBefore, string title, 
        TableOfContentsSwitches switches, 
        string? headerStyle, int maxIncludeLevel, int? rightTabPos)
    {
        ThrowIfNoPackage();

        if (insertBefore == null)
            throw new ArgumentNullException(nameof(insertBefore));

        var toc = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);
        insertBefore.Xml.AddBeforeSelf(toc.Xml);

        return toc;
    }

    /// <summary>
    /// Save this document back to the location it was loaded from.
    /// </summary>
    public void Save()
    {
        ThrowIfNoPackage();

        if (filename == null && stream == null || mainDoc == null || settingsPart == null || memoryStream == null)
        {
            throw DocumentNotLoadedException;
        }

        // Renumber the tracking IDs.
        RenumberIds();

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

        styleManager?.Save();
        numberingStyles?.Save();

        settingsPart.Save(settingsDoc);
        
        if (endnotesDoc != null)
            endnotesPart?.Save(endnotesDoc);
        if (footnotesDoc != null)
            footnotesPart?.Save(footnotesDoc);

        commentManager?.Save();

        // Close the package and commit changes to the memory stream.
        // Note that .NET Core requires we close the package - not just flush it.
        Package.Close();

        // Save back to the file or stream
        if (filename != null)
        {
            memoryStream.Seek(0, SeekOrigin.Begin);
            using var fs = File.Create(filename);
            memoryStream.WriteTo(fs);
        }
        else if (stream != null)
        {
            if (stream.CanSeek)
            {
                stream.SetLength(0);
                stream.Seek(0, SeekOrigin.Begin);
            }

            memoryStream.WriteTo(stream);
        }

        // Reopen the package.
        documentPackage = Package.Open(memoryStream, FileMode.Open, FileAccess.ReadWrite);
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
    /// Adds a new styles.xml to the package and relates it to this document.
    /// </summary>
    private StyleManager AddDefaultStyles()
    {
        if (styleManager == null)
        {
            // If the document contains no /word/styles.xml create one and associate it
            if (!Package.PartExists(Relations.Styles.Uri))
            {
                var stylesDoc = HelperFunctions.AddDefaultStylesXml(Package, out var stylesPart);
                if (stylesDoc.Element(Namespace.Main + "styles") == null)
                    throw new Exception("Missing root styles collection.");

                styleManager = new StyleManager(this, stylesPart);
            }
            else
            {
                styleManager = new StyleManager(this, Package.GetPart(Relations.Styles.Uri));
            }
        }

        return styleManager;
    }

    /// <summary>
    /// Adds the hyperlink style to the document. This is done the first time
    /// a hyperlink is added. If the style already exists, this method does nothing.
    /// </summary>
    internal void AddHyperlinkStyle()
    {
        if (!Styles.HasStyle("Hyperlink", StyleType.Character))
            Styles.Add(Resource.HyperlinkStyle(RevisionId));
    }

    /// <summary>
    /// Recreate the links to the different package parts when we're re-creating the package.
    /// </summary>
    private void RefreshDocumentParts()
    {
        ThrowIfNoPackage();

        // Get the main document part
        var packagePart = Package.GetParts().Single(p =>
            p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase) ||
            p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

        // Set the base DocXElement up.
        SetOwner(this, packagePart, false);

        // Load the settings
        settingsPart = Package.GetPart(Relations.Settings.Uri);

        // Create a comment manager
        commentManager = new CommentManager(this);

        // Load all the sections
        foreach (var rel in PackagePart.GetRelationships())
        {
            if (rel.RelationshipType == Relations.Endnotes.RelType)
                endnotesPart = Package.GetPart(Relations.Endnotes.Uri);
            else if (rel.RelationshipType == Relations.Footnotes.RelType)
                footnotesPart = Package.GetPart(Relations.Footnotes.Uri);
            else if (rel.RelationshipType == Relations.Styles.RelType)
                styleManager = new StyleManager(this, Package.GetPart(Relations.Styles.Uri));
            else if (rel.RelationshipType == Relations.Numbering.RelType)
                numberingStyles = new NumberingStyleManager(this, Package.GetPart(Relations.Numbering.Uri));
            else if (rel.RelationshipType == Relations.People.RelType)
                commentManager.PeoplePackagePart = Package.GetPart(Relations.People.Uri);
            else if (rel.RelationshipType == Relations.Comments.RelType)
                commentManager.CommentsPackagePart = Package.GetPart(Relations.Comments.Uri);
        }
    }

    /// <summary>
    /// Loads the document contents from the assigned package
    /// </summary>
    private void LoadDocumentParts()
    {
        // Load all the package parts
        RefreshDocumentParts();

        // Grab the main document
        mainDoc = PackagePart.Load();

        var bodyElement = mainDoc.Root?.Element(Namespace.Main + "body");

        // Set the DocElement XML value
        Xml = bodyElement ?? throw new DocumentFormatException(nameof(Package), "Main document in package is not properly formed (Body missing).");

        // Get the last revision id
        string? revValue = Xml.GetSectionProperties().RevisionId;
        if (uint.TryParse(revValue, NumberStyles.HexNumber, null, out revision))
            revision++; // bump revision
        else 
            revision = 1;

        // Load all the XML files
        settingsDoc = settingsPart?.Load();
        endnotesDoc = endnotesPart?.Load();
        footnotesDoc = footnotesPart?.Load();
    }

    /// <summary>
    /// This adds the default numbering.xml to the document if it doesn't exist.
    /// </summary>
    private NumberingStyleManager AddDefaultNumberingPart()
    {
        // If we don't have any numbering styles in this document, add a default document.
        if (numberingStyles == null)
        {
            var packagePart = Package.CreatePart(Relations.Numbering.Uri, Relations.Numbering.ContentType, CompressionOption.Maximum);
            var template = Resource.NumberingXml();
            packagePart.Save(template);
            PackagePart.CreateRelationship(packagePart.Uri, TargetMode.Internal, Relations.Numbering.RelType);
            numberingStyles = new NumberingStyleManager(this, packagePart);
        }

        // See if we have the list style
        if (!Styles.HasStyle("ListParagraph", StyleType.Paragraph))
        {
            Styles.Add(Resource.ListParagraphStyle(RevisionId));
        }

        return numberingStyles;
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
    private void ThrowIfNoPackage()
    {
        if (documentPackage == null)
            throw new ObjectDisposedException("Document object has been disposed.");
    }

    /// <summary>
    /// Renumber the tracking ids
    /// </summary>
    private void RenumberIds()
    {
        ThrowIfNoPackage();

        var trackerIds = mainDoc!.Descendants()
            .Where(d => d.Name.LocalName is "ins" or "del")
            .Select(d => d.Attribute(Name.Id))
            .Where(attr => attr != null)
            .Cast<XAttribute>()
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
        ThrowIfNoPackage();

        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));

        string matchPattern = $@"{name.ToUpperInvariant()} \\\* MERGEFORMAT";

        // Update in all sections.
        var documents = new[] {mainDoc!.Root}
            .Union(Sections.SelectMany(s => s.Headers).Select(h => h.Xml))
            .Union(Sections.SelectMany(s => s.Footers).Select(f => f.Xml))
            .Where(element => element != null)
            .Cast<XElement>();

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
        if (settingsDoc == null) return;

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
        ThrowIfNoPackage();
        if (mainDoc == null) throw DocumentNotLoadedException;

        if (nextDocumentId == 0)
            nextDocumentId = HelperFunctions.FindLastUsedDocId(mainDoc);
        return ++nextDocumentId;
    }

    /// <summary>
    /// Creates a new document comment which can be associated to a DocxElement
    /// </summary>
    /// <param name="authorName">Author name</param>
    /// <param name="text">Initial text</param>
    /// <param name="dateTime">Optional date</param>
    /// <param name="authorInitials">Optional initials</param>
    /// <returns>Comment</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public Comment CreateComment(string authorName, string text, DateTime? dateTime, string? authorInitials)
    {
        ThrowIfNoPackage();
        if (commentManager == null) throw DocumentNotLoadedException;

        if (text == null) throw new ArgumentNullException(nameof(text));
        if (string.IsNullOrWhiteSpace(authorName))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(authorName));
            
        var comment = commentManager.CreateComment(authorName, dateTime, authorInitials);
        comment.AddParagraph(text);
        return comment;
    }

    /// <summary>
    /// Returns the raw XML document
    /// </summary>
    /// <returns></returns>
    public string RawDocument() => Xml.ToString();

    /// <summary>
    /// Create a new Picture from a loaded image relationship.
    /// </summary>
    /// <param name="rid">A unique id that identifies an Image embedded in this document.</param>
    /// <param name="name">The name of this Picture.</param>
    /// <param name="description">The description of this Picture.</param>
    internal Drawing CreateDrawingWithEmbeddedPicture(string rid, string name, string description)
    {
        if (string.IsNullOrWhiteSpace(rid))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(rid));

        long id = GetNextDocumentId();
        var (cx, cy) = GetImageDimensions(rid);
        var image = GetRelatedImage(rid);
        if (image == null) throw new ArgumentException("Relationship not found.", nameof(rid));

        Image? svgImage = null;

        // If this image is an SVG, then we need to create a PNG version
        // of the same image. The PNG is what gets rendered in the document
        // but we can keep the SVG to pull out in source format later.
        if (image.ImageType == ".svg")
        {
            svgImage = image; // original image
            image = CreatePngFromSvg(image); // placeholder
        }

        // Create the XML block to represent the drawing + picture.
        var drawingXml = Resource.DrawingElement(id, name, description, cx, cy, image.Id, image.FileName);
            
        // Create the drawing owner.
        var drawing = new Drawing(this, PackagePart, drawingXml);
            
        // Create the picture.
        var picture = new Picture(this, PackagePart,
            drawingXml.FirstLocalNameDescendant("pic")!, image);

        if (svgImage != null)
        {
            // Override local DPI.
            picture.ImageExtensions.Add(new LocalDpiExtension(false));
            // Add in the SVG extension.
            picture.ImageExtensions.Add(new SvgExtension(this, svgImage.Id));
        }

        return drawing;
    }

    /// <summary>
    /// Generate a placeholder PNG from an SVG image.
    /// </summary>
    /// <param name="image">SVG image object</param>
    /// <returns>PNG image object</returns>
    private Image CreatePngFromSvg(Image image)
    {
        // Load the svg
        var svg = new SkiaSharp.Extended.Svg.SKSvg();
        var relationship = PackagePart.GetRelationship(image.Id);
        SkiaSharp.SKPicture? pict;

        using (var svgStream = Package.GetPart(relationship.TargetUri).GetStream())
        {
            pict = svg.Load(svgStream);
        }

        // Get the dimensions.
        var dimen = new SkiaSharp.SKSizeI(
            (int)Math.Ceiling(pict.CullRect.Width),
            (int)Math.Ceiling(pict.CullRect.Height));
        var matrix = SkiaSharp.SKMatrix.MakeScale(1, 1);
        var img = SkiaSharp.SKImage.FromPicture(pict, dimen, matrix);

        // convert to PNG
        var skdata = img.Encode(SkiaSharp.SKEncodedImageFormat.Png, quality:100);
        using var pngStream = new MemoryStream();
        skdata.SaveTo(pngStream);

        // Create an image + relationship from the stream.
        pngStream.Seek(0, SeekOrigin.Begin);
        return AddImage(pngStream, ImageContentType.Png);
    }

    /// <summary>
    /// Returns the image dimension
    /// </summary>
    /// <param name="rid"></param>
    /// <returns></returns>
    private (int width, int height) GetImageDimensions(string rid)
    {
        if (string.IsNullOrEmpty(rid))
            throw new ArgumentException($"'{nameof(rid)}' cannot be null or empty.", nameof(rid));

        //TODO: find a better way to do this.
        int cx = 100;
        int cy = 100;

        var relationship = PackagePart.GetRelationship(rid);
        using (var partStream = Package.GetPart(relationship.TargetUri).GetStream())
        {
            if (Path.GetExtension(relationship.TargetUri.ToString()).ToLower() == ".svg")
            {
                string svg = new StreamReader(partStream).ReadToEnd();
                int sp = svg.IndexOf("viewBox=\"", StringComparison.Ordinal);
                if (sp > 0)
                {
                    sp += 9;
                    int ep = svg.IndexOf("\"", sp, StringComparison.Ordinal);
                    var values = svg.Substring(sp, ep - sp);
                    var split = values.Split(' ');
                    if (split.Length == 4)
                    {
                        cx = (int) double.Parse(split[2]);
                        cy = (int) double.Parse(split[3]);
                    }
                }
            }
            else
            {
                using var img = System.Drawing.Image.FromStream(partStream);
                cx = img.Width;
                cy = img.Height;
            }
        }

        // Return in EMUs.
        return (cx * (int)Uom.EmuConversion, cy * (int)Uom.EmuConversion);
    }

    /// <summary>
    /// Returns an image object from a relationship ID.
    /// </summary>
    /// <param name="rid">Relationship id</param>
    /// <returns>Image</returns>
    internal Image? GetRelatedImage(string rid)
    {
        if (string.IsNullOrEmpty(rid))
            return null;

        var relationship = PackagePart.GetRelationship(rid);
        if (relationship == null)
            throw new ArgumentException("Missing relationship", nameof(rid));

        return new Image(this, relationship);
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
        if (part is null)
            throw new ArgumentNullException(nameof(part));

        // Convert the path of this mainPart to its equivalent rels file path.
        string path = part.Uri.OriginalString.Replace("/word/", "");
        Uri relationshipPath = new Uri($"/word/_rels/{path}.rels", UriKind.Relative);

        // Check to see if the rels file exists and create it if not.
        if (!Document.Package.PartExists(relationshipPath))
        {
            var pp = Document.Package.CreatePart(relationshipPath, DocxContentType.Relationships, CompressionOption.Maximum);
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
        var partUri = PackagePart.GetRelationship(id).TargetUri;
        if (!partUri.OriginalString.StartsWith("/word/"))
        {
            partUri = new Uri("/word/" + partUri.OriginalString, UriKind.Relative);
        }

        // Get the PackagePart and return the XM document.
        part = Package.GetPart(partUri);

        // We keep the header/footer XML documents in memory so we don't have
        // to save them back to the package every time a change is made.
        if (!headerFooterCache.TryGetValue(id, out var existingDoc))
        {
            doc = part.Load();
            AdjustHeaderFooterCache(id, doc);
        }
        else doc = existingDoc;
    }

    /// <summary>
    /// Add or remove a header/footer to/from the cache
    /// </summary>
    /// <param name="id">Relationship id in the main document</param>
    /// <param name="doc">XML document for the header/footer</param>
    internal void AdjustHeaderFooterCache(string id, XDocument? doc)
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
    /// <returns>FirstParagraph</returns>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    internal Paragraph? FindParagraphByIndex(int index)
    {
        if (index < 0)
            throw new ArgumentOutOfRangeException(nameof(index));

        // Collect all the paragraphs based on what character they end on.
        var lookup = Paragraphs
            .Where(p => p.StartIndex < p.EndIndex)
            .ToDictionary(paragraph => paragraph.EndIndex);

        // If the insertion position is first (0) and there are no paragraphs, then return null.
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

    /// <summary>
    /// Returns a Run object for a specific XML fragment if it's part of the document structure.
    /// </summary>
    /// <param name="xml">Xml element to look for</param>
    /// <returns>Created run object</returns>
    internal Run? FindRunByElement(XElement xml) =>
        Paragraphs.Union(Tables.SelectMany(t => t.Paragraphs))
            .SelectMany(p => p.Runs)
            .FirstOrDefault(r => r.Xml == xml);

    private string? filename;               // The filename that this document was loaded from; can be null;
    private Stream? stream;                 // The stream that this document was loaded from; can be null.
    private MemoryStream? memoryStream;     // The in-memory document (with changes)
    private long nextDocumentId;            // Next document id for pictures, bookmarks, etc.

    /// <summary>
    /// The ZIP package holding this Document structure
    /// The ZIP package holding this Document structure
    /// </summary>
    internal Package Package
    {
        get
        {
            ThrowIfNoPackage();
            return documentPackage!;
        }
    }

    /// <summary>
    /// Comment manager (separate XML document)
    /// </summary>
    internal CommentManager CommentManager
    {
        get
        {
            if (commentManager == null)
                throw new ObjectDisposedException(nameof(CommentManager));
            return commentManager;
        }
    }

    /// <summary>
    /// XML documents representing loaded sections of the DOCX file.
    /// These have possible unsaved edits
    /// </summary>
    private XDocument? footnotesDoc;
    private XDocument? endnotesDoc;
    private XDocument? mainDoc;
    private XDocument? settingsDoc;
    private readonly Dictionary<string, XDocument> headerFooterCache = new();

    // Package (.zip) for the document. This contains multiple "parts" each serialized by 
    // an XML document.
    private Package? documentPackage;

    /// <summary>
    /// Package sections in the above Package object. These are specific read/write points in the ZIP file
    /// </summary>
    private PackagePart? footnotesPart;
    private PackagePart? endnotesPart;
    private PackagePart? settingsPart;

    /// <summary>
    /// Styles associated with this document
    /// </summary>
    private StyleManager? styleManager;

    /// <summary>
    /// Document revision
    /// </summary>
    private uint revision;

    /// <summary>
    /// Numbering styles associated with this document
    /// </summary>
    private NumberingStyleManager? numberingStyles;

    /// <summary>
    /// Comment manager for the document
    /// </summary>
    private CommentManager? commentManager;

    /// <summary>
    /// Exception thrown if we end up in a bad state
    /// </summary>
    private static readonly Exception DocumentNotLoadedException = new InvalidOperationException("Document not loaded.");
}