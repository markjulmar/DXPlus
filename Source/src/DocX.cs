using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DXPlus.Charts;
using DXPlus.Helpers;

namespace DXPlus
{
	/// <summary>
	/// Represents a document.
	/// </summary>
	public sealed class DocX : Container, IDisposable
	{
		private string filename;               // The filename that this document was loaded from; can be null;
		private Stream stream;                 // The stream that this document was loaded from; can be null.
		private MemoryStream memoryStream;     // The in-memory document (with changes)

		/// <summary>
		/// The ZIP package holding this DocX structure
		/// </summary>
		internal Package Package { get; set; }

		/// <summary>
		/// XML documents representing loaded sections of the DOCX file.
		/// These have possible unsaved edits
		/// </summary>
		private XDocument fontTableDoc;
		private XDocument footnotesDoc;
		private XDocument endnotesDoc;
		private XDocument stylesWithEffectsDoc;
		internal XDocument mainDoc;
		internal XDocument numberingDoc;
		internal XDocument settingsDoc;
		internal XDocument stylesDoc;

		/// <summary>
		/// Package sections in the above Package object. These are specific read/write points in the ZIP file
		/// </summary>
		private PackagePart fontTablePart;
		private PackagePart footnotesPart;
		private PackagePart numberingPart;
		private PackagePart endnotesPart;
		private PackagePart settingsPart;
		private PackagePart stylesPart;
		private PackagePart stylesWithEffectsPart;
		private string editingSession;

		// A lookup for the Paragraphs in this document
		internal Dictionary<int, Paragraph> paragraphLookup = new Dictionary<int, Paragraph>();

		/// <summary>
		/// Default constructor
		/// </summary>
		internal DocX() : base(null, null)
		{
		}

		/// <summary>
		/// Editing session id for this session.
		/// </summary>
        public string EditingSessionId => editingSession ??= HelperFunctions.GenLongHexNumber();

        /// <summary>
        /// Retrieve all bookmarks
        /// </summary>
        public BookmarkCollection Bookmarks => new BookmarkCollection(Paragraphs.SelectMany(p => p.GetBookmarks()));

		///<summary>
		/// Returns the list of document core properties with corresponding values.
		///</summary>
		public Dictionary<string, string> CoreProperties
		{
			get
			{
				ThrowIfObjectDisposed();
				if (!Package.PartExists(DocxSections.DocPropsCoreUri)) 
					return new Dictionary<string, string>();

				// Get all of the core properties in this document
				var corePropDoc = Package.GetPart(DocxSections.DocPropsCoreUri).Load();
				return corePropDoc.Root.Elements()
					.Select(docProperty =>
						new KeyValuePair<string, string>(
							$"{corePropDoc.Root.GetPrefixOfNamespace(docProperty.Name.Namespace)}:{docProperty.Name.LocalName}",
							docProperty.Value))
					.ToDictionary(p => p.Key, v => v.Value);

			}
		}

		/// <summary>
		/// Returns whether the given style exists in the style catalog
		/// </summary>
		/// <param name="styleId"></param>
		/// <param name="type"></param>
		/// <returns></returns>
		public bool HasStyle(string styleId, string type)
		{
			ThrowIfObjectDisposed();

			return stylesDoc.Descendants(DocxNamespace.Main + "style").Any(x =>
					(x.Attribute(DocxNamespace.Main + "type")?.Value.Equals(type) != false)
					&& x.Attribute(DocxNamespace.Main + "styleId")?.Value.Equals(styleId) == true);
		}

		/// <summary>
		/// Returns a list of custom properties in this document.
		/// </summary>
		public Dictionary<string, CustomProperty> CustomProperties
		{
			get
			{
				ThrowIfObjectDisposed();

				if (Package.PartExists(DocxSections.DocPropsCustom))
				{
					var customPropDoc = Package.GetPart(DocxSections.DocPropsCustom).Load();

					// Get all of the custom properties in this document
					return (
						from p in customPropDoc.Descendants(DocxNamespace.CustomPropertiesSchema + "property")
						let Name = p.AttributeValue(DocxNamespace.Main + "name")
						let Type = p.Descendants().Single().Name.LocalName
						let Value = p.Descendants().Single().Value
						select new CustomProperty(Name, Type, Value)
					).ToDictionary(p => p.Name, StringComparer.CurrentCultureIgnoreCase);
				}

				return new Dictionary<string, CustomProperty>();
			}
		}

		private XElement SectPr
		{
			get
			{
				ThrowIfObjectDisposed();
				return mainDoc.Root!.Element(DocxNamespace.Main + "body")
					.GetOrCreateElement(DocxNamespace.Main + "sectPr");
			}
		}

		/// <summary>
		/// Should the Document use an independent Header and Footer for the first page?
		/// </summary>
		public bool DifferentFirstPage
		{
			get => SectPr.Element(DocxNamespace.Main + "titlePg") != null;

			set
			{
				var titlePg = SectPr.Element(DocxNamespace.Main + "titlePg");
				if (titlePg == null && value)
				{
					SectPr.Add(new XElement(DocxNamespace.Main + "titlePg", string.Empty));
				}
				else if (titlePg != null && !value)
				{
					titlePg.Remove();
				}
			}
		}

		/// <summary>
		/// Should the Document use different Headers and Footers for odd and even pages?
		/// </summary>
		public bool DifferentOddAndEvenPages
		{
			get
			{
				ThrowIfObjectDisposed();
				return settingsDoc.Root.Element(DocxNamespace.Main + "evenAndOddHeaders") != null;
			}

			set
			{
				ThrowIfObjectDisposed();
				var evenAndOddHeaders = settingsDoc.Root.Element(DocxNamespace.Main + "evenAndOddHeaders");
				if (evenAndOddHeaders == null && value)
				{
					settingsDoc.Root.AddFirst(new XElement(DocxNamespace.Main + "evenAndOddHeaders"));
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
		public IEnumerable<string> EndnotesText =>
			endnotesDoc.Root?.Elements(DocxNamespace.Main + "endnote")
				.Select(HelperFunctions.GetText);

		/// <summary>
		/// Returns a collection of Footers in this Document.
		/// A document typically contains three Footers.
		/// A default one (odd), one for the first page and one for even pages.
		/// </summary>
		public Footers Footers { get; private set; }

		/// <summary>
		/// Get the text of each footnote from this document
		/// </summary>
		public IEnumerable<string> FootnotesText =>
			footnotesDoc.Root?.Elements(DocxNamespace.Main + "footnote")
				.Select(HelperFunctions.GetText);

		/// <summary>
		/// Returns a collection of Headers in this Document.
		/// A document typically contains three Headers.
		/// A default one (odd), one for the first page and one for even pages.
		/// </summary>
		public Headers Headers { get; private set; }

		/// <summary>
		/// Returns a list of Images in this document.
		/// </summary>
		public List<Image> Images
		{
			get
			{
				ThrowIfObjectDisposed();
				var imageRelationships = PackagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				return imageRelationships.Any()
					? imageRelationships.Select(i => new Image(this, i)).ToList()
					: new List<Image>();
			}
		}

		/// <summary>
		/// Bottom margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public double MarginBottom
		{
			get => GetMarginAttribute(DocxNamespace.Main + "bottom");
			set => SetMarginAttribute(DocxNamespace.Main + "bottom", value);
		}

		/// <summary>
		/// Left margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public double MarginLeft
		{
			get => GetMarginAttribute(DocxNamespace.Main + "left");
			set => SetMarginAttribute(DocxNamespace.Main + "left", value);
		}

		/// <summary>
		/// Right margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public double MarginRight
		{
			get => GetMarginAttribute(DocxNamespace.Main + "right");
			set => SetMarginAttribute(DocxNamespace.Main + "right", value);
		}

		/// <summary>
		/// Top margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public double MarginTop
		{
			get => GetMarginAttribute(DocxNamespace.Main + "top");
			set => SetMarginAttribute(DocxNamespace.Main + "top", value);
		}

		/// <summary>
		/// Page height value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public double PageHeight
		{
			get
			{
				var pgSz = SectPr.Element(DocxNamespace.Main + "pgSz");
				var w = pgSz?.Attribute(DocxNamespace.Main + "h");
				return w != null && double.TryParse(w.Value, out double value) ? Math.Round(value / 20.0) : 15840.0 / 20.0;
			}

			set => SectPr.GetOrCreateElement(DocxNamespace.Main + "pgSz")
						 .SetAttributeValue(DocxNamespace.Main + "h", value * 20);
		}

		public PageLayout PageLayout => new PageLayout(this, SectPr);

		/// <summary>
		/// Page width value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public double PageWidth
		{
			get
			{
				var pgSz = SectPr.Element(DocxNamespace.Main + "pgSz");
				var w = pgSz?.Attribute(DocxNamespace.Main + "w");
				return w != null && double.TryParse(w.Value, out double f) ? Math.Round(f / 20.0) : 12240.0 / 20.0;
			}

			set => SectPr.Element(DocxNamespace.Main + "pgSz")?
					  .SetAttributeValue(DocxNamespace.Main + "w", value * 20.0);
		}

		/// <summary>
		/// Get the Text of this document.
		/// </summary>
		public string Text => HelperFunctions.GetText(Xml);

		/// <summary>
		/// Creates a document using a fully qualified or relative filename.
		/// </summary>
		/// <param name="filename">The fully qualified or relative filename.</param>
		/// <param name="documentType"></param>
		/// <returns>Returns a DocX object which represents the document.</returns>
		public static DocX Create(string filename, DocumentTypes documentType = DocumentTypes.Document)
		{
			// Create the docx package
			using var ms = new MemoryStream();
			using var package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);
			PostCreation(package, documentType);

			// Load into a document
			var document = Load(ms);
			document.filename = filename;
			document.stream = null;

			// Grab the editing ID we generated.
			document.editingSession = document.Xml.Element(DocxNamespace.Main + "sectPr")
												  .AttributeValue(DocxNamespace.Main + "rsidR");

			return document;
		}

		/// <summary>
		/// Loads a document into a DocX object using a Stream.
		/// </summary>
		/// <param name="stream">The Stream to load the document from.</param>
		/// <returns>
		/// Returns a DocX object which represents the document.
		/// </returns>
		public static DocX Load(Stream stream)
		{
			if (stream == null)
				throw new ArgumentNullException(nameof(stream));
			
			var ms = new MemoryStream();
			stream.Seek(0, SeekOrigin.Begin);
			stream.CopyTo(ms);

			// Open the docx package
			var package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

			var document = PostLoad(package);
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
		public static DocX Load(string filename)
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

			var document = PostLoad(package);
			document.filename = filename;
			document.memoryStream = ms;

			return document;
		}

		/// <summary>
		/// Add a core property to this document. If a core property already exists with the same name it will be replaced. Core property names are case insensitive.
		/// </summary>
		///<param name="propertyName">The property name.</param>
		///<param name="propertyValue">The property value.</param>
		public void AddCoreProperty(string propertyName, string propertyValue)
		{
			ThrowIfObjectDisposed();

			if (string.IsNullOrWhiteSpace(propertyName))
				throw new ArgumentException("Value cannot be null or whitespace.", nameof(propertyName));
			if (string.IsNullOrWhiteSpace(propertyValue))
				throw new ArgumentException("Value cannot be null or whitespace.", nameof(propertyValue));
			if (!Package.PartExists(DocxSections.DocPropsCoreUri))
				throw new Exception("Core properties part doesn't exist.");

			string propertyNamespacePrefix = propertyName.Contains(":") ? propertyName.Split(new[] { ':' })[0] : "cp";
			string propertyLocalName = propertyName.Contains(":") ? propertyName.Split(new[] { ':' })[1] : propertyName;

			var corePropPart = Package.GetPart(DocxSections.DocPropsCoreUri);
			var corePropDoc = corePropPart.Load();

			var corePropElement = corePropDoc.Root.Elements().SingleOrDefault(e => e.Name.LocalName.Equals(propertyLocalName));
			if (corePropElement != null)
			{
				corePropElement.SetValue(propertyValue);
			}
			else
			{
				var propertyNamespace = corePropDoc.Root.GetNamespaceOfPrefix(propertyNamespacePrefix);
				corePropDoc.Root.Add(new XElement(DocxNamespace.Main + propertyLocalName, propertyNamespace.NamespaceName, propertyValue));
			}

			corePropPart.Save(corePropDoc);
			UpdateCorePropertyValue(this, propertyLocalName, propertyValue);
		}

		/// <summary>
		/// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
		/// </summary>
		/// <param name="cp">The CustomProperty to add to this document.</param>
		public void AddCustomProperty(CustomProperty cp)
		{
			ThrowIfObjectDisposed();

			PackagePart customPropertiesPart;
			XDocument customPropDoc;

			// If this document does not contain a custom properties section create one.
			if (!Package.PartExists(DocxSections.DocPropsCustom))
			{
				customPropertiesPart = Package.CreatePart(new Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml", CompressionOption.Maximum);
				customPropDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
					new XElement(DocxNamespace.CustomPropertiesSchema + "Properties",
						new XAttribute(XNamespace.Xmlns + "vt", DocxNamespace.CustomVTypesSchema)
					)
				);

				customPropertiesPart.Save(customPropDoc);
				Package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, $"{DocxNamespace.RelatedDoc.NamespaceName}/custom-properties");
			}
			else
			{
				customPropertiesPart = Package.GetPart(DocxSections.DocPropsCustom);
				customPropDoc = customPropertiesPart.Load();
			}

			// Get the next property id in the document
			var pid = customPropDoc.LocalNameDescendants("property")
				.Select(p => int.TryParse(p.AttributeValue(DocxNamespace.Main + "pid"), out int result) ? result : 0)
				.DefaultIfEmpty().Max() + 1;

			// Check if a custom property already exists with this name - if so, remove it.
			customPropDoc.LocalNameDescendants("property")
					.SingleOrDefault(p => p.AttributeValue(DocxNamespace.Main + "name")
					.Equals(cp.Name, StringComparison.InvariantCultureIgnoreCase))
					?.Remove();

			var propertiesElement = customPropDoc.Element(DocxNamespace.CustomPropertiesSchema + "Properties");
			propertiesElement.Add(
				new XElement(DocxNamespace.CustomPropertiesSchema + "property",
					new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
					new XAttribute("pid", pid),
					new XAttribute("name", cp.Name),
						new XElement(DocxNamespace.CustomVTypesSchema + cp.Type, cp.Value ?? string.Empty)
				)
			);

			// Save the custom properties
			customPropertiesPart.Save(customPropDoc);

			// Refresh all fields in this document which display this custom property.
			UpdateCustomPropertyValue(this, cp.Name, (cp.Value ?? string.Empty).ToString());
		}

		/// <summary>
		/// Adds three new Headers to this document. One for the first page, one for odd pages and one for even pages.
		/// </summary>
		public void AddHeaders()
		{
			AddHeadersOrFooters(true);

			Headers.Odd = Document.GetHeaderByType("default");
			Headers.Even = Document.GetHeaderByType("even");
			Headers.First = Document.GetHeaderByType("first");
		}

		/// <summary>
		/// Adds three new Footers to this document. One for the first page, one for odd pages and one for even pages.
		/// </summary>
		public void AddFooters()
		{
			AddHeadersOrFooters(false);

			Footers.Odd = Document.GetFooterByType("default");
			Footers.Even = Document.GetFooterByType("even");
			Footers.First = Document.GetFooterByType("first");
		}

		/// <summary>
		/// Remove headers from this document
		/// </summary>
		public void RemoveHeaders()
		{
			DeleteHeadersOrFooters(true);
			Headers.Even = Headers.Odd = Headers.First = null;
		}

		/// <summary>
		/// Remove headers from this document
		/// </summary>
		public void RemoveFooters()
		{
			DeleteHeadersOrFooters(false);
			Footers.Even = Footers.Odd = Footers.First = null;
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
			var relationshipsByImages = PackagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
			var imageParts = relationshipsByImages.Select(ir => Package.GetParts()
						.FirstOrDefault(p => p.Uri.ToString().EndsWith(ir.TargetUri.ToString())))
						.Where(e => e != null)
						.ToList();

			foreach (var relsPart in Package.GetParts().Where(part => part.Uri.ToString().Contains("/word/")
																	  && part.ContentType.Equals("application/vnd.openxmlformats-package.relationships+xml")))
			{
				var relsPartContent = relsPart.Load();
				var imageRelationships = relsPartContent.Root.Elements()
											.Where(imageRel => imageRel.Attribute("Type")
													.Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"));

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

				// Compare this image to the new image being added.
				if (HelperFunctions.IsSameFile(existingImageStream, imageStream))
				{
					// Get the image object for this image part
					string id = PackagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
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
			var imageRelation = PackagePart.CreateRelationship(imagePackagePart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

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

		/// <summary>
		/// Returns true if any editing restrictions are imposed on this document.
		/// </summary>
		public bool IsProtected => settingsDoc.Descendants(DocxNamespace.Main + "documentProtection").Any();

		/// <summary>
		/// Add editing protection to this document.
		/// </summary>
		/// <param name="editRestrictions">The type of protection to add to this document.</param>
		public void AddProtection(EditRestrictions editRestrictions)
		{
			AddProtection(editRestrictions, null);
		}

		/// <summary>
		/// Add edit restrictions with a password
		/// http://blogs.msdn.com/b/vsod/archive/2010/04/05/how-to-set-the-editing-restrictions-in-word-using-open-xml-sdk-2-0.aspx
		/// </summary>
		/// <param name="editRestrictions"></param>
		/// <param name="password"></param>
		public void AddProtection(EditRestrictions editRestrictions, string password)
		{
			RemoveProtection();

			if (editRestrictions != EditRestrictions.None)
			{
				var documentProtection = new XElement(DocxNamespace.Main + "documentProtection",
					new XAttribute(DocxNamespace.Main + "edit", editRestrictions.GetEnumName()),
					new XAttribute(DocxNamespace.Main + "enforcement", 1));

				if (!string.IsNullOrWhiteSpace(password))
				{
					int[] initialCodeArray = { 0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3 };
					int[,] encryptionMatrix = new int[15, 7]
					{
						 {0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09}, /* char 1  */
						 {0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF}, /* char 2  */
						 {0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0}, /* char 3  */
						 {0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40}, /* char 4  */
						 {0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5}, /* char 5  */
						 {0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A}, /* char 6  */
						 {0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9}, /* char 7  */
						 {0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0}, /* char 8  */
						 {0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC}, /* char 9  */
						 {0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10}, /* char 10 */
						 {0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168}, /* char 11 */
						 {0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C}, /* char 12 */
						 {0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD}, /* char 13 */
						 {0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC}, /* char 14 */
						 {0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4}  /* char 15 */
					};

					// Generate the Salt
					byte[] arrSalt = new byte[16];
					RandomNumberGenerator rand = new RNGCryptoServiceProvider();
					rand.GetNonZeroBytes(arrSalt);

					const int maxPasswordLength = 15;
					byte[] generatedKey = new byte[4];

					if (!string.IsNullOrEmpty(password))
					{
						password = password.Substring(0, Math.Min(password.Length, maxPasswordLength));

						byte[] arrByteChars = new byte[password.Length];

						for (int intLoop = 0; intLoop < password.Length; intLoop++)
						{
							int intTemp = Convert.ToInt32(password[intLoop]);
							arrByteChars[intLoop] = Convert.ToByte(intTemp & 0x00FF);
							if (arrByteChars[intLoop] == 0)
							{
								arrByteChars[intLoop] = Convert.ToByte((intTemp & 0xFF00) >> 8);
							}
						}

						int intHighOrderWord = initialCodeArray[arrByteChars.Length - 1];

						for (int intLoop = 0; intLoop < arrByteChars.Length; intLoop++)
						{
							int tmp = maxPasswordLength - arrByteChars.Length + intLoop;
							for (int intBit = 0; intBit < 7; intBit++)
							{
								if ((arrByteChars[intLoop] & (0x0001 << intBit)) != 0)
								{
									intHighOrderWord ^= encryptionMatrix[tmp, intBit];
								}
							}
						}

						int intLowOrderWord = 0;

						// For each character in the strPassword, going backwards
						for (int intLoopChar = arrByteChars.Length - 1; intLoopChar >= 0; intLoopChar--)
						{
							intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars[intLoopChar];
						}

						intLowOrderWord = (((intLowOrderWord >> 14) & 0x0001) | ((intLowOrderWord << 1) & 0x7FFF)) ^ arrByteChars.Length ^ 0xCE4B;

						// Combine the Low and High Order Word
						int intCombinedkey = (intHighOrderWord << 16) + intLowOrderWord;

						// The byte order of the result shall be reversed [Example: 0x64CEED7E becomes 7EEDCE64. end example],
						// and that value shall be hashed as defined by the attribute values.
						for (int intTemp = 0; intTemp < 4; intTemp++)
						{
							generatedKey[intTemp] = Convert.ToByte(((uint)(intCombinedkey & (0x000000FF << (intTemp * 8)))) >> (intTemp * 8));
						}
					}

					StringBuilder sb = new StringBuilder();
					for (int intTemp = 0; intTemp < 4; intTemp++)
					{
						sb.Append(Convert.ToString(generatedKey[intTemp], 16));
					}
					generatedKey = Encoding.Unicode.GetBytes(sb.ToString().ToUpper());

					byte[] tmpArray1 = generatedKey;
					byte[] tmpArray2 = arrSalt;
					byte[] tempKey = new byte[tmpArray1.Length + tmpArray2.Length];
					Buffer.BlockCopy(tmpArray2, 0, tempKey, 0, tmpArray2.Length);
					Buffer.BlockCopy(tmpArray1, 0, tempKey, tmpArray2.Length, tmpArray1.Length);
					generatedKey = tempKey;

					const int iterations = 100000;
					HashAlgorithm sha1 = new SHA1Managed();
					generatedKey = sha1.ComputeHash(generatedKey);
					byte[] iterator = new byte[4];
					for (int intTmp = 0; intTmp < iterations; intTmp++)
					{
						iterator[0] = Convert.ToByte((intTmp & 0x000000FF) >> 0);
						iterator[1] = Convert.ToByte((intTmp & 0x0000FF00) >> 8);
						iterator[2] = Convert.ToByte((intTmp & 0x00FF0000) >> 16);
						iterator[3] = Convert.ToByte((intTmp & 0xFF000000) >> 24);

						generatedKey = HelperFunctions.ConcatByteArrays(iterator, generatedKey);
						generatedKey = sha1.ComputeHash(generatedKey);
					}

					documentProtection.Add(new XAttribute(DocxNamespace.Main + "cryptProviderType", "rsaFull"));
					documentProtection.Add(new XAttribute(DocxNamespace.Main + "cryptAlgorithmClass", "hash"));
					documentProtection.Add(new XAttribute(DocxNamespace.Main + "cryptAlgorithmType", "typeAny"));
					documentProtection.Add(new XAttribute(DocxNamespace.Main + "cryptAlgorithmSid", "4"));          // SHA1
					documentProtection.Add(new XAttribute(DocxNamespace.Main + "cryptSpinCount", iterations.ToString()));
					documentProtection.Add(new XAttribute(DocxNamespace.Main + "hash", Convert.ToBase64String(generatedKey)));
					documentProtection.Add(new XAttribute(DocxNamespace.Main + "salt", Convert.ToBase64String(arrSalt)));
				}

				settingsDoc.Root!.AddFirst(documentProtection);
			}
		}

		/// <summary>
		/// Returns the type of editing protection imposed on this document.
		/// </summary>
		/// <returns>The type of editing protection imposed on this document.</returns>
		public EditRestrictions GetProtectionType()
		{
			return IsProtected
				   && settingsDoc.Descendants(DocxNamespace.Main + "documentProtection").First()
					   .Attribute(DocxNamespace.Main + "edit").TryGetEnumValue<EditRestrictions>(out var result)
				? result
				: EditRestrictions.None;
		}

		/// <summary>
		/// Remove editing protection from this document.
		/// </summary>
		public void RemoveProtection()
		{
			ThrowIfObjectDisposed();

			settingsDoc
				.Descendants(DocxNamespace.Main + "documentProtection")
				.Remove();
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
					LoadDocumentParts();
				}
				
				if (!includeContent)
				{
					foreach (var paragraph in Paragraphs)
					{
						paragraph.Remove(false);
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
		/// Saves and copies the document into a new DocX object
		/// </summary>
		/// <returns>
		/// Returns a new DocX object with an identical document
		/// </returns>
		public DocX Copy()
		{
			ThrowIfObjectDisposed();

			var ms = new MemoryStream();
			SaveAs(ms);
			ms.Seek(0, SeekOrigin.Begin);

			return Load(ms);
		}

		/// <summary>
		/// Adds a hyperlink to a document and creates a Paragraph which uses it.
		/// </summary>
		/// <param name="text">The text as displayed by the hyperlink.</param>
		/// <param name="uri">The hyperlink itself.</param>
		/// <returns>Returns a hyperlink that can be inserted into a Paragraph.</returns>
		public Hyperlink CreateHyperlink(string text, Uri uri)
		{
			AddHyperlinkStyleIfNotPresent();

			var data = new XElement(DocxNamespace.Main + "hyperlink",
				new XAttribute(DocxNamespace.RelatedDoc + "id", string.Empty),
				new XAttribute(DocxNamespace.Main + "history", "1"),
				new XElement(DocxNamespace.Main + "r",
					new XElement(DocxNamespace.Main + "rPr",
						new XElement(DocxNamespace.Main + "rStyle",
							new XAttribute(DocxNamespace.Main + "val", "Hyperlink"))),
					new XElement(DocxNamespace.Main + "t", text))
			);

			return new Hyperlink(this, data, PackagePart) { Uri = uri, Text = text };
		}

		/// <summary>
		/// Create a new list tied to this document.
		/// </summary>
		public List CreateList()
		{
			// Ensure we have numbering
			AddDefaultNumberingPart();
			return new List(this, null);
		}

		/// <summary>
		/// Create a new list tied to this document with a specific XML body
		/// </summary>
		/// <param name="xml"></param>
		/// <returns></returns>
		internal List CreateListFromXml(XElement xml)
		{
			// Ensure we have numbering
			AddDefaultNumberingPart();
			return new List(this, xml);
		}

		/// <summary>
		/// Create a new table
		/// </summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		public Table CreateTable(int rowCount, int columnCount)
		{
			if (rowCount < 1)
				throw new ArgumentOutOfRangeException(nameof(rowCount), "Cannot be less than one.");
			if (columnCount < 1)
				throw new ArgumentOutOfRangeException(nameof(columnCount), "Cannot be less than one.");

			return new Table(this, HelperFunctions.CreateTable(rowCount, columnCount)) { PackagePart = PackagePart };
		}

		/// <summary>
		/// Releases all resources used by this document.
		/// </summary>
		public void Dispose()
		{
			(Package as IDisposable)?.Dispose();
			Package = null;
			memoryStream?.Dispose();
			memoryStream = null;
			filename = null;
			stream = null;
		}

		/// <summary>
		/// Retrieve the sections of the document
		/// </summary>
		/// <returns>List of sections</returns>
		public List<Section> GetSections()
		{
			ThrowIfObjectDisposed();

			var parasInASection = new List<Paragraph>();
			var sections = new List<Section>();

			foreach (var para in Paragraphs)
			{
				var sectionInPara = para.Xml.Descendants().FirstOrDefault(s => s.Name.LocalName == "sectPr");
				if (sectionInPara == null)
				{
					parasInASection.Add(para);
				}
				else
				{
					parasInASection.Add(para);
					sections.Add(new Section(Document, sectionInPara) { SectionParagraphs = parasInASection });
					
					parasInASection = new List<Paragraph>();
				}
			}

			// Add the final paragraph
			sections.Add(new Section(Document, SectPr) { SectionParagraphs = parasInASection });

			return sections;
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
			_ = PackagePart.CreateRelationship(chartPackagePart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", id);

			// Save a chart info the chartPackagePart
			chartPackagePart.Save(chart.Xml);

			// Insert a new chart into a paragraph.
			var p = InsertParagraph();
			var chartElement = new XElement(DocxNamespace.Main + "r",
				new XElement(DocxNamespace.Main + "drawing",
					new XElement(DocxNamespace.WordProcessingDrawing + "inline",
						new XElement(DocxNamespace.WordProcessingDrawing + "extent",
							new XAttribute("cx", 5486400),
							new XAttribute("cy", 3200400)),
						new XElement(DocxNamespace.WordProcessingDrawing + "effectExtent",
							new XAttribute("l", 0),
							new XAttribute("t", 0),
							new XAttribute("r", 19050),
							new XAttribute("b", 19050)),
						new XElement(DocxNamespace.WordProcessingDrawing + "docPr",
							new XAttribute("id", 1),
							new XAttribute("name", "chart")),
						new XElement(DocxNamespace.DrawingMain + "graphic",
							new XElement(DocxNamespace.DrawingMain + "graphicData",
								new XAttribute("uri", DocxNamespace.Chart.NamespaceName),
								new XElement(DocxNamespace.Chart + "chart",
									new XAttribute(DocxNamespace.RelatedDoc + "id", id)
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
			return InsertTableOfContents("Table of contents", TableOfContentsSwitches.O | TableOfContentsSwitches.H | TableOfContentsSwitches.Z | TableOfContentsSwitches.U);
		}

		/// <summary>
		/// Insert the contents of another document at the end of this document.
		/// If the document being inserted contains Images, CustomProperties and or custom styles, these will be correctly inserted into the new document.
		/// In the case of Images, new ID's are generated for the Images being inserted to avoid ID conflicts. CustomProperties with the same name will be ignored not replaced.
		/// </summary>
		/// <param name="otherDoc">The document to insert at the end of this document.</param>
		/// <param name="append">If true, document is inserted at the end, otherwise document is inserted at the beginning.</param>
		public void InsertDocument(DocX otherDoc, bool append = true)
		{
			if (otherDoc == null)
				throw new ArgumentNullException(nameof(otherDoc));
			
			ThrowIfObjectDisposed();
			otherDoc.ThrowIfObjectDisposed();

			// Copy all the XML bits
			var otherMainDoc = new XDocument(otherDoc.mainDoc);
			var otherFootnotes = otherDoc.footnotesDoc != null ? new XDocument(otherDoc.footnotesDoc) : null;
			var otherEndnotes = otherDoc.endnotesDoc != null ? new XDocument(otherDoc.endnotesDoc) : null;
			var otherBody = otherMainDoc.Root.Element(DocxNamespace.Main + "body");

			// Remove all header and footer references.
			otherMainDoc.Descendants(DocxNamespace.Main + "headerReference").Remove();
			otherMainDoc.Descendants(DocxNamespace.Main + "footerReference").Remove();

			// Set of content types to ignore
			var ignoreContentTypes = new List<string>
			{
				"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
				"application/vnd.openxmlformats-package.core-properties+xml",
				"application/vnd.openxmlformats-officedocument.extended-properties+xml",
				"application/vnd.openxmlformats-package.relationships+xml",
			};

			// Valid image content types
			var imageContentTypes = new List<string>
			{
				"image/jpeg",
				"image/jpg",
				"image/png",
				"image/bmp",
				"image/gif",
				"image/tiff",
				"image/icon",
				"image/pcx",
				"image/emf",
				"image/wmf",
				"image/svg"
			};
			
			// Check if each PackagePart pp exists in this document.
			foreach (PackagePart otherPackagePart in otherDoc.Package.GetParts())
			{
				if (ignoreContentTypes.Contains(otherPackagePart.ContentType) || imageContentTypes.Contains(otherPackagePart.ContentType))
					continue;

				// If this external PackagePart already exits then we must merge them.
				if (Package.PartExists(otherPackagePart.Uri))
				{
					PackagePart localPackagePart = Package.GetPart(otherPackagePart.Uri);
					switch (otherPackagePart.ContentType)
					{
						case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
							MergeCustoms(otherPackagePart, localPackagePart);
							break;

						// Merge footnotes (and endnotes) before merging styles, then set the remote_footnotes to the just updated footnotes
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
							MergeFootnotes(otherMainDoc, otherFootnotes);
							otherFootnotes = footnotesDoc;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
							MergeEndnotes(otherMainDoc, otherEndnotes);
							otherEndnotes = endnotesDoc;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
						case "application/vnd.ms-word.stylesWithEffects+xml":
							MergeStyles(otherMainDoc, otherDoc, otherFootnotes, otherEndnotes);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
							MergeFonts(otherDoc);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
							MergeNumbering(otherMainDoc, otherDoc);
							break;
					}
				}

				// If this external PackagePart does not exits in the internal document then we can simply copy it.
				else
				{
					var packagePart = ClonePackagePart(otherPackagePart);
					switch (otherPackagePart.ContentType)
					{
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
							endnotesPart = packagePart;
							endnotesDoc = otherEndnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
							footnotesPart = packagePart;
							footnotesDoc = otherFootnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
							stylesPart = packagePart;
							stylesDoc = stylesPart.Load();
							break;

						case "application/vnd.ms-word.stylesWithEffects+xml":
							stylesWithEffectsPart = packagePart;
							stylesWithEffectsDoc = stylesWithEffectsPart.Load();
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
							fontTablePart = packagePart;
							fontTableDoc = fontTablePart.Load();
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
							numberingPart = packagePart;
							numberingDoc = numberingPart.Load();
							break;
					}

					ClonePackageRelationship(otherDoc, otherPackagePart, otherMainDoc);
				}
			}

			// Convert hyperlink ids over
			foreach (var rel in otherDoc.PackagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"))
			{
				string oldId = rel.Id;
				string newId = PackagePart.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType).Id;
				
				foreach (var hyperlink in otherMainDoc.Descendants(DocxNamespace.Main + "hyperlink"))
				{
					var attrId = hyperlink.Attribute(DocxNamespace.RelatedDoc + "id");
					if (attrId != null && attrId.Value == oldId)
						attrId.SetValue(newId);
				}
			}

			// OLE object links
			foreach (var rel in otherDoc.PackagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"))
			{
				string oldId = rel.Id;
				string newId = PackagePart.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType).Id;
				foreach (var oleObject in otherMainDoc.Descendants(XName.Get("OLEObject", "urn:schemas-microsoft-com:office:office")))
				{
					var attrId = oleObject.Attribute(DocxNamespace.RelatedDoc + "id");
					if (attrId != null && attrId.Value == oldId)
						attrId.SetValue(newId);
				}
			}

			// Images
			foreach (var otherPackageParts in otherDoc.Package.GetParts().Where(otherPackageParts => imageContentTypes.Contains(otherPackageParts.ContentType)))
			{
				MergeImages(otherPackageParts, otherDoc, otherMainDoc, otherPackageParts.ContentType);
			}

			// Get the largest drawing object non-visual property id
			int id = mainDoc.Root.Descendants(DocxNamespace.WordProcessingDrawing + "docPr")
							.Max(e => int.TryParse(e.AttributeValue("id"), out int value) ? value : 0) + 1;
				
			// Bump the ids in the other document so they are all above the local document
			foreach (var docPr in otherBody.Descendants(DocxNamespace.WordProcessingDrawing + "docPr"))
				docPr.SetAttributeValue("id", id++);

			// Add the remote documents contents to this document.
			var localBody = mainDoc.Root.Element(DocxNamespace.Main + "body");
			if (append)
				localBody.Add(otherBody.Elements());
			else
				localBody.AddFirst(otherBody.Elements());

			// Copy any missing root attributes to the local document.
			foreach (var attr in otherMainDoc.Root.Attributes())
			{
				if (mainDoc.Root.Attribute(attr.Name) == null)
					mainDoc.Root.SetAttributeValue(attr.Name, attr.Value);
			}
		}

		/// <summary>
		/// Inserts a TOC into the current document.
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
			Xml.Add(toc.Xml);
			
			return toc;
		}

		/// <summary>
		/// Inserts at TOC into the current document before the provided <paramref name="reference"/>
		/// </summary>
		/// <param name="reference">The paragraph to use as reference</param>
		/// <param name="title">The title of the TOC</param>
		/// <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
		/// <param name="headerStyle">Lets you set the style name of the TOC header</param>
		/// <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
		/// <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
		/// <returns>The inserted TableOfContents</returns>
		public TableOfContents InsertTableOfContents(Paragraph reference, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null)
		{
			ThrowIfObjectDisposed();

			var toc = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);
			reference.Xml.AddBeforeSelf(toc.Xml);
			
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

			var xdecl = new XDeclaration("1.0", "UTF-8", "yes");

			// Save the main document
			PackagePart.Save(mainDoc);

			settingsDoc = settingsPart.Load();

			// Save all the sections
			if (Headers?.Even != null)
			{
				Package.GetPart(PackUriHelper.ResolvePartUri(PackagePart.Uri, PackagePart.GetRelationship(Headers.Even.Id).TargetUri))
					   .Save(new XDocument(xdecl, Headers.Even.Xml));
			}

			if (Headers?.Odd != null)
			{
				Package.GetPart(PackUriHelper.ResolvePartUri(PackagePart.Uri, PackagePart.GetRelationship(Headers.Odd.Id).TargetUri))
					   .Save(new XDocument(xdecl, Headers.Odd.Xml));
			}

			if (Headers?.First != null)
			{
				Package.GetPart(PackUriHelper.ResolvePartUri(PackagePart.Uri, PackagePart.GetRelationship(Headers.First.Id).TargetUri))
					   .Save(new XDocument(xdecl, Headers.First.Xml));
			}

			if (Footers?.Even != null)
			{
				Package.GetPart(PackUriHelper.ResolvePartUri(PackagePart.Uri, PackagePart.GetRelationship(Footers.Even.Id).TargetUri))
					.Save(new XDocument(xdecl, Footers.Even.Xml));
			}

			if (Footers?.Odd != null)
			{
				Package.GetPart(PackUriHelper.ResolvePartUri(PackagePart.Uri, PackagePart.GetRelationship(Footers.Odd.Id).TargetUri))
					   .Save(new XDocument(xdecl, Footers.Odd.Xml));
			}

			if (Footers?.First != null)
			{
				Package.GetPart(PackUriHelper.ResolvePartUri(PackagePart.Uri, PackagePart.GetRelationship(Footers.First.Id).TargetUri))
					   .Save(new XDocument(xdecl, Footers.First.Xml));
			}

			settingsPart.Save(settingsDoc);
			endnotesPart?.Save(endnotesDoc);
			footnotesPart?.Save(footnotesDoc);
			stylesPart?.Save(stylesDoc);
			stylesWithEffectsPart?.Save(stylesWithEffectsDoc);
			numberingPart?.Save(numberingDoc);
			fontTablePart?.Save(fontTableDoc);

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
		/// Load all the internal properties from the created package
		/// </summary>
		/// <param name="package">Document package</param>
		/// <param name="documentType">Document type</param>
		internal static void PostCreation(Package package, DocumentTypes documentType = DocumentTypes.Document)
		{
			if (package == null)
				throw new ArgumentNullException(nameof(package));

			// Create the main document part for this package
			var mainDocumentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative),
				documentType == DocumentTypes.Document ? DocxContentType.Document : DocxContentType.Template,
				CompressionOption.Normal);
			package.CreateRelationship(mainDocumentPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");

			// Generate an id for this editing session.
			var rsid = HelperFunctions.GenLongHexNumber();

			// Load the document part into a XDocument object
			var mainDoc = XDocument.Parse(
			$@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
			   <w:document xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
			   <w:body>
				<w:sectPr w:rsidR=""{rsid}"" w:rsidSect=""{rsid}"">
					<w:pgSz w:h=""15840"" w:w=""12240""/>
					<w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""720"" w:footer=""720"" w:gutter=""0""/>
					<w:cols w:space=""720""/>
					<w:docGrid w:linePitch=""360""/>
				</w:sectPr>
			   </w:body>
			   </w:document>"
			);

			// Add the settings.xml
			HelperFunctions.AddDefaultSettingsPart(package, rsid);

			// Add the default styles
			HelperFunctions.AddDefaultStylesXml(package);

			// Save the main document
			mainDocumentPart.Save(mainDoc);
			package.Close();
		}

		/// <summary>
		/// Load the properties from the loaded package
		/// </summary>
		/// <param name="package">Loaded package</param>
		/// <returns>New Docx object</returns>
		internal static DocX PostLoad(Package package)
		{
			var document = new DocX { Package = package };
			document.Document = document;
			document.LoadDocumentParts();
			
			return document;
		}

		internal static void UpdateCorePropertyValue(DocX document, string name, string value)
		{
			if (document == null)
				throw new ArgumentNullException(nameof(document));
			if (string.IsNullOrWhiteSpace(name))
				throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));

			document.ThrowIfObjectDisposed();

			string matchPattern = $@"(DOCPROPERTY)?{name}\\\*MERGEFORMAT".ToLower();
			
			foreach (var e in document.mainDoc.Descendants(DocxNamespace.Main + "fldSimple"))
			{
				string attrValue = e.AttributeValue(DocxNamespace.Main + "instr").Replace(" ", string.Empty).Trim().ToLower();

				if (Regex.IsMatch(attrValue, matchPattern))
				{
					var firstRun = e.Element(DocxNamespace.Main + "r");
					var rPr = firstRun?.GetRunProps(false);

					e.RemoveNodes();

					if (firstRun != null)
					{
						e.Add(new XElement(firstRun.Name,
                                firstRun.Attributes(),
                                rPr,
								new XElement(DocxNamespace.Main + "t", value).PreserveSpace()
                            )
                        );
					}
				}
			}

			void ProcessHeaderFooterParts(IEnumerable<PackagePart> packageParts)
			{
				foreach (var pp in packageParts)
				{
					var section = pp.Load();

					foreach (var e in section.Descendants(DocxNamespace.Main + "fldSimple"))
					{
						string attrValue = e.AttributeValue(DocxNamespace.Main + "instr").Replace(" ", string.Empty).Trim().ToLower();
						if (Regex.IsMatch(attrValue, matchPattern))
						{
							var firstRun = e.Element(DocxNamespace.Main + "r");
							e.RemoveNodes();
							if (firstRun != null)
							{
								e.Add(new XElement(firstRun.Name, 
                                        firstRun.Attributes(), 
                                        firstRun.GetRunProps(false), 
										new XElement(DocxNamespace.Main + "t", value).PreserveSpace()
                                    )
                                );
							}
						}
					}

					pp.Save(section);
				}
			}


			ProcessHeaderFooterParts(document.Package.GetParts()
				.Where(headerPart => Regex.IsMatch(headerPart.Uri.ToString(), @"/word/header\d?.xml")));

			ProcessHeaderFooterParts(document.Package.GetParts()
				.Where(footerPart => (Regex.IsMatch(footerPart.Uri.ToString(), @"/word/footer\d?.xml"))));
		}

		/// <summary>
		/// Update the custom properties inside the document
		/// </summary>
		/// <param name="document">The DocX document</param>
		/// <param name="name">The property used inside the document</param>
		/// <param name="value">The new value for the property</param>
		/// <remarks>Different version of Word create different Document XML.</remarks>
		internal static void UpdateCustomPropertyValue(DocX document, string name, string value)
		{
			if (document == null)
				throw new ArgumentNullException(nameof(document));
			if (string.IsNullOrWhiteSpace(name))
				throw new ArgumentNullException(nameof(name));

			var documents = new List<XElement> { document.mainDoc.Root };

			var headers = document.Headers;
			if (headers.First != null)
				documents.Add(headers.First.Xml);

			if (headers.Odd != null)
				documents.Add(headers.Odd.Xml);

			if (headers.Even != null)
				documents.Add(headers.Even.Xml);

			var footers = document.Footers;
			if (footers.First != null)
				documents.Add(footers.First.Xml);

			if (footers.Odd != null)
				documents.Add(footers.Odd.Xml);

			if (footers.Even != null)
				documents.Add(footers.Even.Xml);

			string matchCustomPropertyName = name;
			if (name.Contains(" "))
				matchCustomPropertyName = "\"" + name + "\"";

			string propertyMatchValue = $@"DOCPROPERTY  {matchCustomPropertyName}  \* MERGEFORMAT".Replace(" ", string.Empty);

			// Process each document in the list.
			foreach (var doc in documents)
			{
				foreach (var e in doc.Descendants(DocxNamespace.Main + "instrText"))
				{
					string attrValue = e.Value.Replace(" ", string.Empty).Trim();

					if (attrValue.Equals(propertyMatchValue, StringComparison.CurrentCultureIgnoreCase))
					{
						var nextNode = e.Parent.NextNode;
						bool found = false;
						while (true)
						{
							if (nextNode.NodeType == XmlNodeType.Element)
							{
								var ele = (XElement) nextNode;
								var match = ele.Descendants(DocxNamespace.Main + "t");
								if (match.Any())
								{
									if (!found)
									{
										match.First().Value = value;
										found = true;
									}
									else
									{
										ele.RemoveNodes();
									}
								}
								else
								{
									match = ele.Descendants(DocxNamespace.Main + "fldChar");
									if (match.Any())
									{
										var endMatch = match.First().Attribute(DocxNamespace.Main + "fldCharType");
										if (endMatch?.Value == "end")
											break;
									}
								}
							}
							nextNode = nextNode.NextNode;
						}
					}
				}

				foreach (var e in doc.Descendants(DocxNamespace.Main + "fldSimple"))
				{
					string attrValue = e.Attribute(DocxNamespace.Main + "instr").Value.Replace(" ", string.Empty).Trim();

					if (attrValue.Equals(propertyMatchValue, StringComparison.CurrentCultureIgnoreCase))
					{
						var firstRun = e.Element(DocxNamespace.Main + "r");
						var firstText = firstRun.Element(DocxNamespace.Main + "t");
						var rPr = firstText.GetRunProps(false);

						// Delete everything and insert updated text value
						e.RemoveNodes();

						e.Add(new XElement(firstRun.Name, 
                                firstRun.Attributes(), 
                                rPr, 
								new XElement(DocxNamespace.Main + "t", value).PreserveSpace()
                            )
                        );
					}
				}
			}
		}

		/// <summary>
		/// Adds a Header or Footer to a document.
		/// If the document already contains a Header it will be replaced.
		/// </summary>
		/// <param name="addHeader">True for header, False for footer</param>
		internal void AddHeadersOrFooters(bool addHeader)
		{
			string element = addHeader ? "hdr" : "ftr";
			string reference = addHeader ? "header" : "footer";

			// Delete header/footer
			DeleteHeadersOrFooters(addHeader);

			var rsid = EditingSessionId;

			int index = 1;
			foreach (var headerType in new[] { "default", "even", "first" })
			{
				var headerPart = Package.CreatePart(new Uri($"/word/{reference}{index}.xml", UriKind.Relative),
					$"application/vnd.openxmlformats-officedocument.wordprocessingml.{reference}+xml", CompressionOption.Normal);
				var headerRelationship = PackagePart.CreateRelationship(headerPart.Uri, TargetMode.Internal,
					$"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{reference}");

				var header = XDocument.Parse(
					$@"<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>
				   <w:{element} xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">
					 <w:p w:rsidR=""{rsid}"" w:rsidRDefault=""{rsid}"">
					   <w:pPr>
						 <w:pStyle w:val=""{reference}"" />
					   </w:pPr>
					 </w:p>
				   </w:{element}>"
				);

				// Save the main document
				headerPart.Save(header);

				SectPr.Add(new XElement(DocxNamespace.Main + $"{reference}Reference",
						new XAttribute(DocxNamespace.Main + "type", headerType),
						new XAttribute(DocxNamespace.RelatedDoc + "id", headerRelationship.Id)
					));
				
				index++;
			}
		}

		internal void AddHyperlinkStyleIfNotPresent()
		{
			ThrowIfObjectDisposed();

			// If the document contains no /word/styles.xml create one and associate it
			if (!Package.PartExists(DocxSections.StylesUri))
			{
				HelperFunctions.AddDefaultStylesXml(Package);
			}

			// Ensure we are looking at the correct one.
			stylesPart ??= Package.GetPart(DocxSections.StylesUri);
			stylesDoc ??= stylesPart.Load();

			// Check for the "hyperlinkStyle"
			bool hyperlinkStyleExists = stylesDoc.Element(DocxNamespace.Main + "styles").Elements()
				.Any(e => e.Attribute(DocxNamespace.Main + "styleId")?.Value == "Hyperlink");

			if (!hyperlinkStyleExists)
			{
				// Add a simple Hyperlink style (blue + underline + default font + size)
				stylesDoc.Element(DocxNamespace.Main + "styles").Add(
					new XElement
					(
						DocxNamespace.Main + "style",
						new XAttribute(DocxNamespace.Main + "type", "character"),
						new XAttribute(DocxNamespace.Main + "styleId", "Hyperlink"),
							new XElement(DocxNamespace.Main + "name", new XAttribute(DocxNamespace.Main + "val", "Hyperlink")),
							new XElement(DocxNamespace.Main + "basedOn", new XAttribute(DocxNamespace.Main + "val", "DefaultParagraphFont")),
							new XElement(DocxNamespace.Main + "uiPriority", new XAttribute(DocxNamespace.Main + "val", "99")),
							new XElement(DocxNamespace.Main + "unhideWhenUsed"),
							new XElement(DocxNamespace.Main + "rsid", new XAttribute(DocxNamespace.Main + "val", "0005416C")),
							new XElement(DocxNamespace.Main + "rPr",
								new XElement(DocxNamespace.Main + "color", new XAttribute(DocxNamespace.Main + "val", "0000FF"), new XAttribute(DocxNamespace.Main + "themeColor", "hyperlink")),
								new XElement(DocxNamespace.Main + "u", new XAttribute(DocxNamespace.Main + "val", "single"))
							)
					)
				);
			}
		}

		internal void DeleteHeadersOrFooters(bool isHeader)
		{
			ThrowIfObjectDisposed();

			string reference = (isHeader) ? "header" : "footer";

			// Get all Header/Footer relationships in this document.
			foreach (var rel in PackagePart.GetRelationshipsByType(
				$"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{reference}"))
			{
				var uri = rel.TargetUri;
				if (!uri.OriginalString.StartsWith("/word/"))
					uri = new Uri("/word/" + uri.OriginalString, UriKind.Relative);

				// Check to see if the document actually contains the Part.
				if (Package.PartExists(uri))
				{
					// Delete the part
					Package.DeletePart(uri);

					// Get all references to this relationship in the document and remove them.
					mainDoc.Descendants(DocxNamespace.Main + "body")
						   .Descendants().Where(e => e.Name.LocalName == $"{reference}Reference" 
								&& e.AttributeValue(DocxNamespace.RelatedDoc + "id") == rel.Id)
						   .Remove();

					// Delete the Relationship.
					Package.DeleteRelationship(rel.Id);
				}
			}
		}

		/// <summary>
		/// Recreate the links to the different package parts when we're re-creating the package.
		/// </summary>
		internal void RefreshDocumentParts()
		{
			ThrowIfObjectDisposed();

			PackagePart = Package.GetParts().Single(p =>
				p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase) ||
				p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

			settingsPart = Package.GetPart(DocxSections.SettingsUri);

			// Load all the optional sections
			foreach (var rel in PackagePart.GetRelationships())
			{
				switch (rel.RelationshipType)
				{
					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes":
						endnotesPart = Package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes":
						footnotesPart = Package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
						stylesPart = Package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						break;

					case "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects":
						stylesWithEffectsPart = Package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
						fontTablePart = Package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numberingDoc":
						numberingPart = Package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						break;
				}
			}
		}

		internal void LoadDocumentParts()
		{
			if (Package == null)
				throw new InvalidOperationException($"Cannot load package parts when {nameof(Package)} property is not set.");
			
			// Load all the package parts
			RefreshDocumentParts();

			// Grab the main document
			mainDoc = PackagePart.Load();

			// Set the DocElement XML value
			Xml = mainDoc.Root!.Element(DocxNamespace.Main + "body");
			if (Xml == null)
				throw new InvalidOperationException($"Missing {DocxNamespace.Main + "body"} expected content.");

			// Load headers
			Headers = new Headers {
				Odd = GetHeaderByType("default"),
				Even = GetHeaderByType("even"),
				First = GetHeaderByType("first")
			};
			
			// Load footers
			Footers = new Footers {
				Odd = GetFooterByType("default"),
				Even = GetFooterByType("even"),
				First = GetFooterByType("first")
			};

			// Load all the XML files
			settingsDoc = settingsPart?.Load();
			endnotesDoc = endnotesPart?.Load();
			footnotesDoc = footnotesPart?.Load();
			stylesDoc = stylesPart?.Load();
			stylesWithEffectsDoc = stylesWithEffectsPart?.Load();
			fontTableDoc = fontTablePart?.Load();
			numberingDoc = numberingPart?.Load();

			// Load all the paragraphs
			paragraphLookup.Clear();
			foreach (var paragraph in Paragraphs)
			{
				if (!paragraphLookup.ContainsKey(paragraph.EndIndex))
				{
					paragraphLookup.Add(paragraph.EndIndex, paragraph);
				}
			}
		}

		private void AddDefaultNumberingPart()
		{
			if (numberingPart == null)
			{
				numberingPart = Package.CreatePart(new Uri("/word/numbering.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", CompressionOption.Maximum);
				numberingDoc = Resources.NumberingXml();
				numberingPart.Save(numberingDoc);
				PackagePart.CreateRelationship(numberingPart.Uri, TargetMode.Internal, $"{DocxNamespace.RelatedDoc.NamespaceName}/numbering");
			}

			// If the document contains no /word/styles.xml create one and associate it
			if (!Package.PartExists(DocxSections.StylesUri))
			{
				HelperFunctions.AddDefaultStylesXml(Package);
			}

			// Ensure we are looking at the correct one.
			stylesPart ??= Package.GetPart(DocxSections.StylesUri);
			stylesDoc ??= stylesPart.Load();

			var rootStyleNode = stylesDoc.Element(DocxNamespace.Main + "styles");
			if (rootStyleNode == null)
				throw new InvalidOperationException("Missing root style node after creation.");

			// See if we have the list style
			bool listStyleExists =
			(
				from s in rootStyleNode.Elements()
				let styleId = s.Attribute(DocxNamespace.Main + "styleId")
				where (styleId?.Value == "ListParagraph")
				select s
			).Any();

			if (!listStyleExists)
			{
				var style = new XElement(DocxNamespace.Main + "style",
					new XAttribute(DocxNamespace.Main + "type", "paragraph"),
					new XAttribute(DocxNamespace.Main + "styleId", "ListParagraph"),
					new XElement(DocxNamespace.Main + "name",
						new XAttribute(DocxNamespace.Main + "val", "List Paragraph")),
					new XElement(DocxNamespace.Main + "basedOn",
						new XAttribute(DocxNamespace.Main + "val", "Normal")),
					new XElement(DocxNamespace.Main + "uiPriority",
						new XAttribute(DocxNamespace.Main + "val", "34")),
					new XElement(DocxNamespace.Main + "qformat"),
					new XElement(DocxNamespace.Main + "rsid",
						new XAttribute(DocxNamespace.Main + "val", "00832EE1")),
					new XElement(DocxNamespace.Main + "rPr",
						new XElement(DocxNamespace.Main + "ind",
							new XAttribute(DocxNamespace.Main + "left", "720")),
						new XElement(DocxNamespace.Main + "contextualSpacing")
					)
				);

				rootStyleNode.Add(style);
			}
		}

		/// <summary>
		/// Clone a package part
		/// </summary>
		/// <param name="part">Part to clone</param>
		/// <returns>Clone of passed part</returns>
		private PackagePart ClonePackagePart(PackagePart part)
		{
			if (part == null)
				throw new ArgumentNullException(nameof(part));
			
			ThrowIfObjectDisposed();

			var newPackagePart = Package.CreatePart(part.Uri, part.ContentType, CompressionOption.Normal);

			using var sr = part.GetStream();
			using var sw = newPackagePart.GetStream(FileMode.Create);
			sr.CopyTo(sw);

			return newPackagePart;
		}

		private void ClonePackageRelationship(DocX otherDocument, PackagePart part, XDocument otherXml)
		{
			if (otherDocument == null)
				throw new ArgumentNullException(nameof(otherDocument));
			if (part == null)
				throw new ArgumentNullException(nameof(part));
			if (otherXml == null)
				throw new ArgumentNullException(nameof(otherXml));

			ThrowIfObjectDisposed();
			otherDocument.ThrowIfObjectDisposed();

			string url = part.Uri.OriginalString.Replace("/", "");
			foreach (PackageRelationship remoteRelationship in otherDocument.PackagePart.GetRelationships())
			{
				if (url.Equals("word" + remoteRelationship.TargetUri.OriginalString.Replace("/", "")))
				{
					string remoteId = remoteRelationship.Id;
					string localId = PackagePart.CreateRelationship(remoteRelationship.TargetUri, remoteRelationship.TargetMode, remoteRelationship.RelationshipType).Id;

					// Replace all instances of remote id in the local document with the local id
					foreach (var elem in otherXml.Descendants(DocxNamespace.DrawingMain + "blip"))
					{
						var embed = elem.Attribute(DocxNamespace.RelatedDoc + "embed");
						if (embed?.Value == remoteId)
						{
							embed.SetValue(localId);
						}
					}

					// Do the same for shapes
					foreach (var elem in otherXml.Descendants(DocxNamespace.VML + "imagedata"))
					{
						var id = elem.Attribute(DocxNamespace.RelatedDoc + "id");
						if (id?.Value == remoteId)
						{
							id.SetValue(localId);
						}
					}
					
					break;
				}
			}
		}

		private static string ComputeHashString(Stream stream)
		{
			byte[] hash = MD5.Create().ComputeHash(stream);
			
			var sb = new StringBuilder();
			foreach (var value in hash)
				sb.Append(value.ToString("X2"));

			return sb.ToString();
		}

		private Footer GetFooterByType(string type)
		{
			return (Footer) GetHeaderOrFooterByType(type, false);
		}

		private Header GetHeaderByType(string type)
		{
			return (Header) GetHeaderOrFooterByType(type, true);
		}

		private Container GetHeaderOrFooterByType(string type, bool isHeader)
		{
			ThrowIfObjectDisposed();

			string reference = !isHeader ? "footerReference" : "headerReference";

			string id = mainDoc.Descendants(DocxNamespace.Main + "body").Descendants()
				.Where(e => (e.Name.LocalName == reference) && (e.AttributeValue(DocxNamespace.Main + "type") == type))
				.Select(e => e.AttributeValue(DocxNamespace.RelatedDoc + "id"))
				.LastOrDefault();

			if (id == null)
				return null;

			// Get the Xml file for this Header or Footer.
			Uri partUri = PackagePart.GetRelationship(id).TargetUri;

			// Weird problem with PackagePart API.
			if (!partUri.OriginalString.StartsWith("/word/"))
				partUri = new Uri("/word/" + partUri.OriginalString, UriKind.Relative);

			// Get the PackagePart and load the XML
			var part = Package.GetPart(partUri);
			var doc = part.Load();

			// Header and Footer extend Container.
			return isHeader
				? new Header(this, doc.Element(DocxNamespace.Main + "hdr"), part) { Id = id }
				: (Container)new Footer(this, doc.Element(DocxNamespace.Main + "ftr"), part) { Id = id };

		}

		/// <summary>
		/// Get a margin
		/// </summary>
		/// <param name="name">Margin to get</param>
		/// <returns>Value in 1/20th pt.</returns>
		private double GetMarginAttribute(XName name)
		{
			ThrowIfObjectDisposed();

			var top = SectPr.Element(DocxNamespace.Main + "pgMar")?.Attribute(name);
			return top != null && double.TryParse(top.Value, out double value) ? (int) (value / 20.0) : 0;
		}

		/// <summary>
		/// Set a margin
		/// </summary>
		/// <param name="name">Margin to set</param>
		/// <param name="value">Value in 1/20th pt</param>
		private void SetMarginAttribute(XName name, double value)
		{
			ThrowIfObjectDisposed();

			SectPr.GetOrCreateElement(DocxNamespace.Main + "pgMar")
				  .SetAttributeValue(name, value * 20.0);
		}

		private string GetNextRelationshipId()
		{
			// Last used id (0 if none)
			int id = PackagePart.GetRelationships()
						.Where(r => r.Id.Substring(0, 3).Equals("rId"))
						.Select(r => int.TryParse(r.Id.Substring(3), out int result) ? result : 0)
						.DefaultIfEmpty()
						.Max();
			return $"rId{id+1}";
		}

		private void MergeImages(PackagePart remote_pp, DocX remote_document, XDocument remote_mainDoc, string contentType)
		{
			if (remote_pp == null)
				throw new ArgumentNullException(nameof(remote_pp));
			if (remote_document == null)
				throw new ArgumentNullException(nameof(remote_document));
			if (remote_mainDoc == null)
				throw new ArgumentNullException(nameof(remote_mainDoc));
			if (string.IsNullOrWhiteSpace(contentType))
				throw new ArgumentNullException(nameof(contentType));

			// Before doing any other work, check to see if this image is actually referenced in the document.
			PackageRelationship remote_rel = remote_document.PackagePart.GetRelationships().Where(r => r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString.Replace("/word/", ""))).FirstOrDefault();
			if (remote_rel == null)
			{
				remote_rel = remote_document.PackagePart.GetRelationships().Where(r => r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString)).FirstOrDefault();
				if (remote_rel == null)
				{
					return;
				}
			}

			string remote_Id = remote_rel.Id;
			string remote_hash = ComputeHashString(remote_pp.GetStream());
			IEnumerable<PackagePart> image_parts = Package.GetParts().Where(pp => pp.ContentType.Equals(contentType));

			bool found = false;
			foreach (PackagePart part in image_parts)
			{
				string local_hash = ComputeHashString(part.GetStream());
				if (local_hash.Equals(remote_hash))
				{
					// This image already exists in this document.
					found = true;

					PackageRelationship local_rel = PackagePart.GetRelationships().FirstOrDefault(r => r.TargetUri.OriginalString.Equals(part.Uri.OriginalString.Replace("/word/", "")))
								 ?? PackagePart.GetRelationships().FirstOrDefault(r => r.TargetUri.OriginalString.Equals(part.Uri.OriginalString));
					if (local_rel != null)
					{
						string new_Id = local_rel.Id;

						// Replace all instances of remote_Id in the local document with local_Id
						foreach (XElement elem in remote_mainDoc.Descendants(DocxNamespace.DrawingMain + "blip"))
						{
							XAttribute embed = elem.Attribute(DocxNamespace.RelatedDoc + "embed");
							if (embed != null && embed.Value == remote_Id)
							{
								embed.SetValue(new_Id);
							}
						}

						// Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
						foreach (XElement elem in remote_mainDoc.Descendants(DocxNamespace.VML + "imagedata"))
						{
							XAttribute id = elem.Attribute(DocxNamespace.RelatedDoc + "id");
							if (id != null && id.Value == remote_Id)
							{
								id.SetValue(new_Id);
							}
						}
					}

					break;
				}
			}

			// This image does not exist in this document.
			if (!found)
			{
				string new_uri = remote_pp.Uri.OriginalString;
				new_uri = new_uri.Remove(new_uri.LastIndexOf("/"));
				new_uri += "/" + Guid.NewGuid().ToString() + contentType.Replace("image/", ".");
				if (!new_uri.StartsWith("/"))
				{
					new_uri = "/" + new_uri;
				}

				PackagePart new_pp = Package.CreatePart(new Uri(new_uri, UriKind.Relative), remote_pp.ContentType, CompressionOption.Normal);

				using (Stream s_read = remote_pp.GetStream())
				{
					using Stream s_write = new_pp.GetStream(FileMode.Create);
					byte[] buffer = new byte[short.MaxValue];
					int read;
					while ((read = s_read.Read(buffer, 0, buffer.Length)) > 0)
					{
						s_write.Write(buffer, 0, read);
					}
				}

				PackageRelationship pr = PackagePart.CreateRelationship(new Uri(new_uri, UriKind.Relative), TargetMode.Internal,
													"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

				string new_Id = pr.Id;

				//Check if the remote relationship id is a default rId from Word
				Match defRelId = Regex.Match(remote_Id, @"rId\d+", RegexOptions.IgnoreCase);

				// Replace all instances of remote_Id in the local document with local_Id
				foreach (XElement elem in remote_mainDoc.Descendants(DocxNamespace.DrawingMain + "blip"))
				{
					XAttribute embed = elem.Attribute(DocxNamespace.RelatedDoc + "embed");
					if (embed != null && embed.Value == remote_Id)
					{
						embed.SetValue(new_Id);
					}
				}

				if (!defRelId.Success)
				{
					// Replace all instances of remote_Id in the local document with local_Id
					foreach (XElement elem in mainDoc.Descendants(DocxNamespace.DrawingMain + "blip"))
					{
						XAttribute embed = elem.Attribute(DocxNamespace.RelatedDoc + "embed");
						if (embed != null && embed.Value == remote_Id)
						{
							embed.SetValue(new_Id);
						}
					}

					// Replace all instances of remote_Id in the local document with local_Id
					foreach (XElement elem in mainDoc.Descendants(DocxNamespace.VML + "imagedata"))
					{
						XAttribute id = elem.Attribute(DocxNamespace.RelatedDoc + "id");
						if (id != null && id.Value == remote_Id)
						{
							id.SetValue(new_Id);
						}
					}
				}

				// Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
				foreach (XElement elem in remote_mainDoc.Descendants(DocxNamespace.VML + "imagedata"))
				{
					XAttribute id = elem.Attribute(DocxNamespace.RelatedDoc + "id");
					if (id != null && id.Value == remote_Id)
					{
						id.SetValue(new_Id);
					}
				}
			}
		}

		private void MergeCustoms(PackagePart remote_pp, PackagePart local_pp)
		{
			if (remote_pp == null)
				throw new ArgumentNullException(nameof(remote_pp));
			if (local_pp == null)
				throw new ArgumentNullException(nameof(local_pp));

			// Get the remote documents custom.xml file.
			XDocument remote_custom_document;
			using (TextReader tr = new StreamReader(remote_pp.GetStream()))
			{
				remote_custom_document = XDocument.Load(tr);
			}

			// Get the local documents custom.xml file.
			XDocument local_custom_document;
			using (TextReader tr = new StreamReader(local_pp.GetStream()))
			{
				local_custom_document = XDocument.Load(tr);
			}

			IEnumerable<int> pids = remote_custom_document.Root.Descendants()
				.Where(d => d.Name.LocalName == "property")
				.Select(d => int.Parse(d.Attribute("pid").Value));

			int pid = pids.Max() + 1;

			foreach (XElement remote_property in remote_custom_document.Root.Elements())
			{
				bool found = false;
				foreach (XElement local_property in local_custom_document.Root.Elements())
				{
					XAttribute remote_property_name = remote_property.Attribute("name");
					XAttribute local_property_name = local_property.Attribute("name");

					if (remote_property != null && local_property_name != null && remote_property_name.Value.Equals(local_property_name.Value))
					{
						found = true;
					}
				}

				if (!found)
				{
					remote_property.SetAttributeValue("pid", pid);
					local_custom_document.Root.Add(remote_property);
					pid++;
				}
			}

			// Save the modified local custom styles.xml file.
			local_pp.Save(local_custom_document);
		}

		private void MergeEndnotes(XDocument remote_mainDoc, XDocument remote_endnotes)
		{
			if (remote_mainDoc == null)
				throw new ArgumentNullException(nameof(remote_mainDoc));
			if (remote_endnotes == null)
				throw new ArgumentNullException(nameof(remote_endnotes));
			
			IEnumerable<int> ids = endnotesDoc.Root.Descendants()
				.Where(d => d.Name.LocalName == "endnote")
				.Select(d => int.Parse(d.Attribute(DocxNamespace.Main + "id").Value));

			int max_id = ids.Max() + 1;
			IEnumerable<XElement> endnoteReferences = remote_mainDoc.Descendants(DocxNamespace.Main + "endnoteReference");

			foreach (XElement endnote in remote_endnotes.Root.Elements().OrderBy(fr => fr.Attribute(DocxNamespace.RelatedDoc + "id")).Reverse())
			{
				XAttribute id = endnote.Attribute(DocxNamespace.Main + "id");
				if (id != null && int.TryParse(id.Value, out int i) && i > 0)
				{
					foreach (XElement endnoteRef in endnoteReferences)
					{
						XAttribute a = endnoteRef.Attribute(DocxNamespace.Main + "id");
						if (a != null && int.Parse(a.Value).Equals(i))
						{
							a.SetValue(max_id);
						}
					}

					// We care about copying this footnote.
					endnote.SetAttributeValue(DocxNamespace.Main + "id", max_id);
					endnotesDoc.Root.Add(endnote);
					max_id++;
				}
			}
		}

		private void MergeFonts(DocX remote)
		{
			if (remote == null)
				throw new ArgumentNullException(nameof(remote));
			
			// Add each remote font to this document.
			IEnumerable<XElement> remote_fonts = remote.fontTableDoc.Root.Elements(DocxNamespace.Main + "font");
			IEnumerable<XElement> local_fonts = fontTableDoc.Root.Elements(DocxNamespace.Main + "font");

			foreach (XElement remote_font in remote_fonts)
			{
				bool flag_addFont = true;
				foreach (XElement local_font in local_fonts)
				{
					if (local_font.Attribute(DocxNamespace.Main + "name").Value == remote_font.Attribute(DocxNamespace.Main + "name").Value)
					{
						flag_addFont = false;
						break;
					}
				}

				if (flag_addFont)
				{
					fontTableDoc.Root.Add(remote_font);
				}
			}
		}

		private void MergeFootnotes(XDocument remote_mainDoc, XDocument remote_footnotes)
		{
			if (remote_mainDoc == null)
				throw new ArgumentNullException(nameof(remote_mainDoc));
			if (remote_footnotes == null)
				throw new ArgumentNullException(nameof(remote_footnotes));
			
			IEnumerable<int> ids = footnotesDoc.Root.Descendants()
				.Where(d => d.Name.LocalName == "footnote")
				.Select(d => int.Parse(d.Attribute(DocxNamespace.Main + "id").Value));

			int max_id = ids.Max() + 1;
			IEnumerable<XElement> footnoteReferences = remote_mainDoc.Descendants(DocxNamespace.Main + "footnoteReference");

			foreach (XElement footnote in remote_footnotes.Root.Elements().OrderBy(fr => fr.Attribute(DocxNamespace.RelatedDoc + "id")).Reverse())
			{
				XAttribute id = footnote.Attribute(DocxNamespace.Main + "id");
				if (id != null && int.TryParse(id.Value, out int i) && i > 0)
				{
					foreach (XElement footnoteRef in footnoteReferences)
					{
						XAttribute a = footnoteRef.Attribute(DocxNamespace.Main + "id");
						if (a != null && int.Parse(a.Value).Equals(i))
						{
							a.SetValue(max_id);
						}
					}

					// We care about copying this footnote.
					footnote.SetAttributeValue(DocxNamespace.Main + "id", max_id);
					footnotesDoc.Root.Add(footnote);
					max_id++;
				}
			}
		}

		private void MergeNumbering(XDocument remote_mainDoc, DocX remote)
		{
			if (remote_mainDoc == null)
				throw new ArgumentNullException(nameof(remote_mainDoc));
			if (remote == null)
				throw new ArgumentNullException(nameof(remote));
			
			// Add each remote numberingDoc to this document.
			List<XElement> remote_abstractNums = remote.numberingDoc.Root.Elements(DocxNamespace.Main + "abstractNum").ToList();
			int guidd = 0;
			foreach (XElement an in remote_abstractNums)
			{
				XAttribute a = an.Attribute(DocxNamespace.Main + "abstractNumId");
				if (a != null && int.TryParse(a.Value, out int i) && i > guidd)
				{
					guidd = i;
				}
			}
			guidd++;

			List<XElement> remote_nums = remote.numberingDoc.Root.Elements(DocxNamespace.Main + "num").ToList();
			int guidd2 = 0;
			foreach (XElement an in remote_nums)
			{
				XAttribute a = an.Attribute(DocxNamespace.Main + "numId");
				if (a != null && int.TryParse(a.Value, out int i) && i > guidd2)
				{
					guidd2 = i;
				}
			}
			guidd2++;

			foreach (XElement remote_abstractNum in remote_abstractNums)
			{
				XAttribute abstractNumId = remote_abstractNum.Attribute(DocxNamespace.Main + "abstractNumId");
				if (abstractNumId != null)
				{
					string abstractNumIdValue = abstractNumId.Value;
					abstractNumId.SetValue(guidd);

					foreach (XElement remote_num in remote_nums)
					{
						foreach (XElement numId in remote_mainDoc.Descendants(DocxNamespace.Main + "numId"))
						{
							XAttribute attr = numId.Attribute(DocxNamespace.Main + "val");
							if (attr?.Value.Equals(remote_num.Attribute(DocxNamespace.Main + "numId").Value) == true)
							{
								attr.SetValue(guidd2);
							}
						}
						remote_num.SetAttributeValue(DocxNamespace.Main + "numId", guidd2);

						XElement e = remote_num.Element(DocxNamespace.Main + "abstractNumId");
						if (e != null)
						{
							XAttribute a2 = e.Attribute(DocxNamespace.Main + "val");
							if (a2?.Value.Equals(abstractNumIdValue) == true)
							{
								a2.SetValue(guidd);
							}
						}

						guidd2++;
					}
				}

				guidd++;
			}

			// Checking whether there were more than 0 elements, helped me get rid of exceptions thrown while using InsertDocument
			if (numberingDoc?.Root.Elements(DocxNamespace.Main + "abstractNum").Any() == true)
			{
				numberingDoc.Root.Elements(DocxNamespace.Main + "abstractNum").Last().AddAfterSelf(remote_abstractNums);
			}

			if (numberingDoc?.Root.Elements(DocxNamespace.Main + "num").Any() == true)
			{
				numberingDoc.Root.Elements(DocxNamespace.Main + "num").Last().AddAfterSelf(remote_nums);
			}
		}

		private void MergeStyles(XDocument remote_mainDoc, DocX remote, XDocument remote_footnotes, XDocument remote_endnotes)
		{
			if (remote_mainDoc == null)
				throw new ArgumentNullException(nameof(remote_mainDoc));
			if (remote == null)
				throw new ArgumentNullException(nameof(remote));
			if (remote_footnotes == null)
				throw new ArgumentNullException(nameof(remote_footnotes));
			if (remote_endnotes == null)
				throw new ArgumentNullException(nameof(remote_endnotes));
			
			Dictionary<string, string> local_styles = new Dictionary<string, string>();
			foreach (XElement local_style in stylesDoc.Root.Elements(DocxNamespace.Main + "style"))
			{
				XElement temp = new XElement(local_style);
				XAttribute styleId = temp.Attribute(DocxNamespace.Main + "styleId");
				string value = styleId.Value;
				styleId.Remove();
				string key = Regex.Replace(temp.ToString(), @"\s+", "");
				if (!local_styles.ContainsKey(key))
				{
					local_styles.Add(key, value);
				}
			}

			// Add each remote style to this document.
			foreach (XElement remote_style in remote.stylesDoc.Root.Elements(DocxNamespace.Main + "style"))
			{
				XElement temp = new XElement(remote_style);
				XAttribute styleId = temp.Attribute(DocxNamespace.Main + "styleId");
				string value = styleId.Value;
				styleId.Remove();
				string key = Regex.Replace(temp.ToString(), @"\s+", "");
				string guuid;

				// Check to see if the local document already contains the remote style.
				if (local_styles.ContainsKey(key))
				{
					local_styles.TryGetValue(key, out string local_value);

					// If the styleIds are the same then nothing needs to be done.
					if (local_value == value)
					{
						continue;
					}

					// All we need to do is update the styleId.
					else
					{
						guuid = local_value;
					}
				}
				else
				{
					guuid = Guid.NewGuid().ToString();
					// Set the styleId in the remote_style to this new Guid
					// [Fixed the issue that my document referred to a new Guid while my styles still had the old value ("Titel")]
					remote_style.SetAttributeValue(DocxNamespace.Main + "styleId", guuid);
				}

				foreach (XElement e in remote_mainDoc.Root.Descendants(DocxNamespace.Main + "pStyle"))
				{
					XAttribute e_styleId = e.Attribute(DocxNamespace.Main + "val");
					if (e_styleId?.Value.Equals(styleId.Value) == true)
					{
						e_styleId.SetValue(guuid);
					}
				}

				foreach (XElement e in remote_mainDoc.Root.Descendants(DocxNamespace.Main + "rStyle"))
				{
					XAttribute e_styleId = e.Attribute(DocxNamespace.Main + "val");
					if (e_styleId?.Value.Equals(styleId.Value) == true)
					{
						e_styleId.SetValue(guuid);
					}
				}

				foreach (XElement e in remote_mainDoc.Root.Descendants(DocxNamespace.Main + "tblStyle"))
				{
					XAttribute e_styleId = e.Attribute(DocxNamespace.Main + "val");
					if (e_styleId?.Value.Equals(styleId.Value) == true)
					{
						e_styleId.SetValue(guuid);
					}
				}

				if (remote_endnotes != null)
				{
					foreach (XElement e in remote_endnotes.Root.Descendants(DocxNamespace.Main + "rStyle"))
					{
						XAttribute e_styleId = e.Attribute(DocxNamespace.Main + "val");
						if (e_styleId?.Value.Equals(styleId.Value) == true)
						{
							e_styleId.SetValue(guuid);
						}
					}

					foreach (XElement e in remote_endnotes.Root.Descendants(DocxNamespace.Main + "pStyle"))
					{
						XAttribute e_styleId = e.Attribute(DocxNamespace.Main + "val");
						if (e_styleId?.Value.Equals(styleId.Value) == true)
						{
							e_styleId.SetValue(guuid);
						}
					}
				}

				if (remote_footnotes != null)
				{
					foreach (XElement e in remote_footnotes.Root.Descendants(DocxNamespace.Main + "rStyle"))
					{
						XAttribute e_styleId = e.Attribute(DocxNamespace.Main + "val");
						if (e_styleId?.Value.Equals(styleId.Value) == true)
						{
							e_styleId.SetValue(guuid);
						}
					}

					foreach (XElement e in remote_footnotes.Root.Descendants(DocxNamespace.Main + "pStyle"))
					{
						XAttribute e_styleId = e.Attribute(DocxNamespace.Main + "val");
						if (e_styleId?.Value.Equals(styleId.Value) == true)
						{
							e_styleId.SetValue(guuid);
						}
					}
				}

				// Make sure they don't clash by using a uuid.
				styleId.SetValue(guuid);
				stylesDoc.Root.Add(remote_style);
			}
		}

		/// <summary>
		/// Method to throw an ObjectDisposedException
		/// </summary>
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		void ThrowIfObjectDisposed()
		{
			if (Package == null)
				throw new ObjectDisposedException("DocX object has been disposed.");
		}
	}
}