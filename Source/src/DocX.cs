using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace DXPlus
{
	/// <summary>
	/// Represents a document.
	/// </summary>
	public sealed class DocX : Container, IDisposable
	{
		internal string filename;               // The filename that this document was loaded from; can be null;
		internal Stream stream;                 // The stream that this document was loaded from; can be null.

		// Bits of the Word document
		internal Package package;

		internal XDocument fontTable;
		internal PackagePart fontTablePart;
		internal XDocument footnotes;
		internal PackagePart footnotesPart;
		internal XDocument mainDoc;
		internal MemoryStream memoryStream;
		internal XDocument numbering;
		internal PackagePart numberingPart;
		internal XDocument endnotes;
		internal PackagePart endnotesPart;
		internal XDocument settings;
		internal PackagePart settingsPart;
		internal XDocument styles;
		internal PackagePart stylesPart;
		internal XDocument stylesWithEffects;
		internal PackagePart stylesWithEffectsPart;

		// A lookup for the Paragraphs in this document.
		internal Dictionary<int, Paragraph> paragraphLookup = new Dictionary<int, Paragraph>();

		// Keys in the XML Word format
		private const string Text_Body = "body";

		private const string Text_Bottom = "bottom";
		private const string Text_DocumentProtection = "documentProtection";
		private const string Text_Left = "left";
		private const string Text_PageMargins = "pgMar";
		private const string Text_PageSize = "pgSz";
		private const string Text_PageSizeHeight = "h";
		private const string Text_PageSizeWidth = "w";
		private const string Text_Right = "right";
		private const string Text_SectionProperties = "sectPr";
		private const string Text_Top = "top";

		internal DocX() : base(null, null)
		{
		}

		internal DocX(DocX document, XElement xml) : base(document, xml)
		{
		}

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
				if (package.PartExists(DocxSections.DocPropsCoreUri))
				{
					PackagePart docProps_Core = package.GetPart(DocxSections.DocPropsCoreUri);
					XDocument corePropDoc;
					using (TextReader tr = new StreamReader(docProps_Core.GetStream(FileMode.Open, FileAccess.Read)))
						corePropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

					// Get all of the core properties in this document
					return (from docProperty in corePropDoc.Root.Elements()
							select new KeyValuePair<string, string>(
							  $"{corePropDoc.Root.GetPrefixOfNamespace(docProperty.Name.Namespace)}:{docProperty.Name.LocalName}",
							  docProperty.Value)).ToDictionary(p => p.Key, v => v.Value);
				}

				return new Dictionary<string, string>();
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
			return styles.Descendants(DocxNamespace.Main + "style").Any(x =>
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
				if (package.PartExists(DocxSections.DocPropsCustom))
				{
					PackagePart docProps_custom = package.GetPart(DocxSections.DocPropsCustom);
					XDocument customPropDoc;
					using (TextReader tr = new StreamReader(docProps_custom.GetStream(FileMode.Open, FileAccess.Read)))
						customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

					// Get all of the custom properties in this document
					return
					(
						from p in customPropDoc.Descendants(DocxNamespace.CustomPropertiesSchema + "property")
						let Name = p.Attribute(DocxNamespace.Main + "name").Value
						let Type = p.Descendants().Single().Name.LocalName
						let Value = p.Descendants().Single().Value
						select new CustomProperty(Name, Type, Value)
					).ToDictionary(p => p.Name, StringComparer.CurrentCultureIgnoreCase);
				}

				return new Dictionary<string, CustomProperty>();
			}
		}

		private XElement SectPr => mainDoc.Root.Element(DocxNamespace.Main + Text_Body).GetOrCreateElement(DocxNamespace.Main + Text_SectionProperties);

		/// <summary>
		/// Should the Document use an independent Header and Footer for the first page?
		/// </summary>
		public bool DifferentFirstPage
		{
			get => SectPr?.Element(DocxNamespace.Main + "titlePg") != null;

			set
			{
				XElement titlePg = SectPr?.Element(DocxNamespace.Main + "titlePg");
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
				XDocument settings;
				using (TextReader tr = new StreamReader(settingsPart.GetStream()))
					settings = XDocument.Load(tr);
				return settings.Root.Element(DocxNamespace.Main + "evenAndOddHeaders") != null;
			}

			set
			{
				XDocument settings;
				using (TextReader tr = new StreamReader(settingsPart.GetStream()))
					settings = XDocument.Load(tr);

				XElement evenAndOddHeaders = settings.Root.Element(DocxNamespace.Main + "evenAndOddHeaders");
				if (evenAndOddHeaders == null && value)
				{
					settings.Root.AddFirst(new XElement(DocxNamespace.Main + "evenAndOddHeaders"));
				}
				else if (!value)
				{
					evenAndOddHeaders?.Remove();
				}

				using TextWriter tw = new StreamWriter(settingsPart.GetStream());
				settings.Save(tw);
			}
		}

		/// <summary>
		/// Get the text of each endnote from this document
		/// </summary>
		public IEnumerable<string> EndnotesText
		{
			get
			{
				foreach (XElement endnote in endnotes.Root.Elements(DocxNamespace.Main + "endnote"))
				{
					yield return HelperFunctions.GetText(endnote);
				}
			}
		}

		/// <summary>
		/// Returns a collection of Footers in this Document.
		/// A document typically contains three Footers.
		/// A default one (odd), one for the first page and one for even pages.
		/// </summary>
		public Footers Footers { get; private set; }

		/// <summary>
		/// Get the text of each footnote from this document
		/// </summary>
		public IEnumerable<string> FootnotesText
		{
			get
			{
				foreach (XElement footnote in footnotes.Root.Elements(DocxNamespace.Main + "footnote"))
				{
					yield return HelperFunctions.GetText(footnote);
				}
			}
		}

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
				PackageRelationshipCollection imageRelationships = packagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				return imageRelationships.Any()
					? imageRelationships.Select(i => new Image(this, i)).ToList()
					: new List<Image>();
			}
		}

		/// <summary>
		/// Returns true if any editing restrictions are imposed on this document.
		/// </summary>
		public bool IsProtected => settings.Descendants(DocxNamespace.Main + Text_DocumentProtection).Any();

		/// <summary>
		/// Bottom margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public float MarginBottom
		{
			get => GetMarginAttribute(DocxNamespace.Main + Text_Bottom);
			set => SetMarginAttribute(DocxNamespace.Main + Text_Bottom, value);
		}

		/// <summary>
		/// Left margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public float MarginLeft
		{
			get => GetMarginAttribute(DocxNamespace.Main + Text_Left);
			set => SetMarginAttribute(DocxNamespace.Main + Text_Left, value);
		}

		/// <summary>
		/// Right margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public float MarginRight
		{
			get => GetMarginAttribute(DocxNamespace.Main + Text_Right);
			set => SetMarginAttribute(DocxNamespace.Main + Text_Right, value);
		}

		/// <summary>
		/// Top margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public float MarginTop
		{
			get => GetMarginAttribute(DocxNamespace.Main + Text_Top);
			set => SetMarginAttribute(DocxNamespace.Main + Text_Top, value);
		}

		/// <summary>
		/// Page height value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public float PageHeight
		{
			get
			{
				var pgSz = SectPr?.Element(DocxNamespace.Main + Text_PageSize);
				if (pgSz != null)
				{
					XAttribute w = pgSz.Attribute(DocxNamespace.Main + Text_PageSizeHeight);
					if (w != null && float.TryParse(w.Value, out float f))
					{
						return (int)(f / 20.0f);
					}
				}

				return 15840.0f / 20.0f;
			}

			set => SectPr.GetOrCreateElement(DocxNamespace.Main + Text_PageSize)
						 .SetAttributeValue(DocxNamespace.Main + Text_PageSizeHeight, value * 20);
		}

		public PageLayout PageLayout => new PageLayout(this, SectPr);

		/// <summary>
		/// Page width value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		/// </summary>
		public float PageWidth
		{
			get
			{
				XElement pgSz = SectPr.Element(DocxNamespace.Main + Text_PageSize);
				if (pgSz != null)
				{
					XAttribute w = pgSz.Attribute(DocxNamespace.Main + Text_PageSizeWidth);
					if (w != null && float.TryParse(w.Value, out float f))
					{
						return (int)(f / 20.0f);
					}
				}

				return 12240.0f / 20.0f;
			}

			set => SectPr.Element(DocxNamespace.Main + Text_PageSize)?
					  .SetAttributeValue(DocxNamespace.Main + Text_PageSizeWidth, value * 20);
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
			using MemoryStream ms = new MemoryStream();
			using Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);
			PostCreation(package, documentType);

			// Load into a document
			DocX document = Load(ms);
			document.filename = filename;
			document.stream = null;

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
			MemoryStream ms = new MemoryStream();
			stream.Seek(0, SeekOrigin.Begin);
			stream.CopyTo(ms);

			// Open the docx package
			Package package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

			DocX document = PostLoad(ref package);
			document.package = package;
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
			if (!File.Exists(filename))
				throw new FileNotFoundException(string.Format("File could not be found {0}", filename));

			// Open the docx package
			MemoryStream ms = new MemoryStream();
			using (var fs = new FileStream(filename, FileMode.Open))
				fs.CopyTo(ms);
			Package package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite);

			DocX document = PostLoad(ref package);
			document.package = package;
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
			string propertyNamespacePrefix = propertyName.Contains(":") ? propertyName.Split(new[] { ':' })[0] : "cp";
			string propertyLocalName = propertyName.Contains(":") ? propertyName.Split(new[] { ':' })[1] : propertyName;

			// If this document does not contain a coreFilePropertyPart create one.)
			if (!package.PartExists(DocxSections.DocPropsCoreUri))
				throw new Exception("Core properties part doesn't exist.");

			XDocument corePropDoc;
			PackagePart corePropPart = package.GetPart(DocxSections.DocPropsCoreUri);
			using (TextReader tr = new StreamReader(corePropPart.GetStream(FileMode.Open, FileAccess.Read)))
			{
				corePropDoc = XDocument.Load(tr);
			}

			XElement corePropElement = corePropDoc.Root.Elements().SingleOrDefault(e => e.Name.LocalName.Equals(propertyLocalName));
			if (corePropElement != null)
			{
				corePropElement.SetValue(propertyValue);
			}
			else
			{
				var propertyNamespace = corePropDoc.Root.GetNamespaceOfPrefix(propertyNamespacePrefix);
				corePropDoc.Root.Add(new XElement(DocxNamespace.Main + propertyLocalName, propertyNamespace.NamespaceName, propertyValue));
			}

			using (TextWriter tw = new StreamWriter(corePropPart.GetStream(FileMode.Create, FileAccess.Write)))
			{
				corePropDoc.Save(tw);
			}

			UpdateCorePropertyValue(this, propertyLocalName, propertyValue);
		}

		/// <summary>
		/// Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
		/// </summary>
		/// <param name="cp">The CustomProperty to add to this document.</param>
		public void AddCustomProperty(CustomProperty cp)
		{
			// If this document does not contain a customFilePropertyPart create one.
			if (!package.PartExists(DocxSections.DocPropsCustom))
				HelperFunctions.CreateCustomPropertiesPart(this);

			XDocument customPropDoc;
			PackagePart customPropPart = package.GetPart(DocxSections.DocPropsCustom);
			using (TextReader tr = new StreamReader(customPropPart.GetStream(FileMode.Open, FileAccess.Read)))
				customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

			// Each custom property has a PID, get the highest PID in this document.
			IEnumerable<int> pids = customPropDoc.LocalNameDescendants("property")
												 .Select(p => int.Parse(p.AttributeValue(DocxNamespace.Main + "pid")));

			int pid = pids.Any() ? pids.Max() : 1;

			// Check if a custom property already exists with this name
			var customProperty = customPropDoc.LocalNameDescendants("property")
											  .SingleOrDefault(p => p.AttributeValue(DocxNamespace.Main + "name")
												.Equals(cp.Name, StringComparison.InvariantCultureIgnoreCase));

			// If a custom property with this name already exists remove it.
			customProperty?.Remove();

			XElement propertiesElement = customPropDoc.Element(DocxNamespace.CustomPropertiesSchema + "Properties");
			propertiesElement.Add(
				new XElement(DocxNamespace.CustomPropertiesSchema + "property",
					new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
					new XAttribute("pid", pid + 1),
					new XAttribute("name", cp.Name),
						new XElement(DocxNamespace.CustomVTypesSchema + cp.Type, cp.Value ?? string.Empty)
				)
			);

			// Save the custom properties
			using (TextWriter tw = new StreamWriter(customPropPart.GetStream(FileMode.Create, FileAccess.Write)))
				customPropDoc.Save(tw, SaveOptions.None);

			// Refresh all fields in this document which display this custom property.
			UpdateCustomPropertyValue(this, cp.Name, (cp.Value ?? string.Empty).ToString());
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
		/// Add an Image into this document from a fully qualified or relative filename.
		/// </summary>
		/// <param name="filename">The fully qualified or relative filename.</param>
		/// <returns>An Image file.</returns>
		public Image AddImage(string filename)
		{
			// The extension this file has will be taken to be its format.
			string contentType = (Path.GetExtension(filename)) switch
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
			return AddImage(filename, contentType);
		}

		/// <summary>
		/// Add an Image into this document from a Stream.
		/// </summary>
		/// <param name="stream">A Stream stream.</param>
		/// <returns>An Image file.</returns>
		public Image AddImage(Stream stream)
		{
			return AddImage(stream as object);
		}

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
				XElement documentProtection = new XElement(DocxNamespace.Main + Text_DocumentProtection);
				documentProtection.Add(new XAttribute(DocxNamespace.Main + "edit", editRestrictions.GetEnumName()));
				documentProtection.Add(new XAttribute(DocxNamespace.Main + "enforcement", 1));
				settings.Root.AddFirst(documentProtection);

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

					const int MaxPasswordLength = 15;
					byte[] generatedKey = new byte[4];

					if (!string.IsNullOrEmpty(password))
					{
						password = password.Substring(0, Math.Min(password.Length, MaxPasswordLength));

						byte[] arrByteChars = new byte[password.Length];

						for (int intLoop = 0; intLoop < password.Length; intLoop++)
						{
							int intTemp = Convert.ToInt32(password[intLoop]);
							arrByteChars[intLoop] = Convert.ToByte(intTemp & 0x00FF);
							if (arrByteChars[intLoop] == 0)
								arrByteChars[intLoop] = Convert.ToByte((intTemp & 0xFF00) >> 8);
						}

						int intHighOrderWord = initialCodeArray[arrByteChars.Length - 1];

						for (int intLoop = 0; intLoop < arrByteChars.Length; intLoop++)
						{
							int tmp = MaxPasswordLength - arrByteChars.Length + intLoop;
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

					settings.Root.AddFirst(documentProtection);
				}
			}
		}

		///<summary>
		/// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		///</summary>
		///<param name="templateFilePath">The path to the document template file.</param>
		///<exception cref="FileNotFoundException">The document template file not found.</exception>
		public void ApplyTemplate(string templateFilePath)
		{
			ApplyTemplate(templateFilePath, true);
		}

		///<summary>
		/// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		///</summary>
		///<param name="templateFilePath">The path to the document template file.</param>
		///<param name="includeContent">Whether to copy the document template text content to document.</param>
		///<exception cref="FileNotFoundException">The document template file not found.</exception>
		public void ApplyTemplate(string templateFilePath, bool includeContent)
		{
			if (!File.Exists(templateFilePath))
			{
				throw new FileNotFoundException($"File could not be found {templateFilePath}");
			}

			using FileStream packageStream = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read);
			ApplyTemplate(packageStream, includeContent);
		}

		///<summary>
		/// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		///</summary>
		///<param name="templateStream">The stream of the document template file.</param>
		public void ApplyTemplate(Stream templateStream)
		{
			ApplyTemplate(templateStream, true);
		}

		///<summary>
		/// Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		///</summary>
		///<param name="templateStream">The stream of the document template file.</param>
		///<param name="includeContent">Whether to copy the document template text content to document.</param>
		public void ApplyTemplate(Stream templateStream, bool includeContent)
		{
			Package templatePackage = Package.Open(templateStream);
			try
			{
				PackagePart documentPart = null;
				XDocument documentDoc = null;
				foreach (PackagePart packagePart in templatePackage.GetParts())
				{
					switch (packagePart.Uri.ToString())
					{
						case "/word/document.xml":
							documentPart = packagePart;
							using (XmlReader xr = XmlReader.Create(packagePart.GetStream(FileMode.Open, FileAccess.Read)))
							{
								documentDoc = XDocument.Load(xr);
							}
							break;

						case "/_rels/.rels":
							if (!this.package.PartExists(packagePart.Uri))
							{
								this.package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption);
							}
							PackagePart globalRelsPart = this.package.GetPart(packagePart.Uri);
							using (StreamReader tr = new StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8))
							{
								using StreamWriter tw = new StreamWriter(globalRelsPart.GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8);
								tw.Write(tr.ReadToEnd());
							}
							break;

						case "/word/_rels/document.xml.rels":
							break;

						default:
							if (!this.package.PartExists(packagePart.Uri))
							{
								this.package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption);
							}
							Encoding packagePartEncoding = Encoding.Default;
							if (packagePart.Uri.ToString().EndsWith(".xml") || packagePart.Uri.ToString().EndsWith(".rels"))
							{
								packagePartEncoding = Encoding.UTF8;
							}
							PackagePart nativePart = this.package.GetPart(packagePart.Uri);
							using (StreamReader tr = new StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read), packagePartEncoding))
							{
								using StreamWriter tw = new StreamWriter(nativePart.GetStream(FileMode.Create, FileAccess.Write), tr.CurrentEncoding);
								tw.Write(tr.ReadToEnd());
							}
							break;
					}
				}
				if (documentPart != null)
				{
					string mainContentType = documentPart.ContentType.Replace("template.main", "document.main");
					if (this.package.PartExists(documentPart.Uri))
					{
						this.package.DeletePart(documentPart.Uri);
					}
					PackagePart documentNewPart = this.package.CreatePart(
					  documentPart.Uri, mainContentType, documentPart.CompressionOption);
					using (XmlWriter xw = XmlWriter.Create(documentNewPart.GetStream(FileMode.Create, FileAccess.Write)))
					{
						documentDoc.WriteTo(xw);
					}
					foreach (PackageRelationship documentPartRel in documentPart.GetRelationships())
					{
						documentNewPart.CreateRelationship(documentPartRel.TargetUri, documentPartRel.TargetMode, documentPartRel.RelationshipType, documentPartRel.Id);
					}

					this.packagePart = documentNewPart;
					this.mainDoc = documentDoc;

					PopulateDocument(this, templatePackage);
					settingsPart = HelperFunctions.CreateOrGetSettingsPart(package);
				}
				if (!includeContent)
				{
					foreach (Paragraph paragraph in this.Paragraphs)
					{
						paragraph.Remove(false);
					}
				}
			}
			finally
			{
				this.package.Flush();
				var documentRelsPart = this.package.GetPart(new Uri("/word/_rels/document.xml.rels", UriKind.Relative));
				using (TextReader tr = new StreamReader(documentRelsPart.GetStream(FileMode.Open, FileAccess.Read)))
				{
					tr.Read();
				}
				templatePackage.Close();
				PopulateDocument(Document, package);
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
			MemoryStream ms = new MemoryStream();
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

			XElement data = new XElement(DocxNamespace.Main + "hyperlink",
				new XAttribute(DocxNamespace.RelatedDoc + "id", string.Empty),
				new XAttribute(DocxNamespace.Main + "history", "1"),
				new XElement(DocxNamespace.Main + "r",
					new XElement(DocxNamespace.Main + "rPr",
						new XElement(DocxNamespace.Main + "rStyle",
							new XAttribute(DocxNamespace.Main + "val", "Hyperlink"))),
					new XElement(DocxNamespace.Main + "t", text))
			);

			return new Hyperlink(this, data, uri);
		}

		/// <summary>
		/// Create a new list tied to this document.
		/// </summary>
		public List CreateList() => new List(this, null);

		/// <summary>
		/// Create a new table
		/// </summary>
		/// <param name="rowCount"></param>
		/// <param name="columnCount"></param>
		public Table CreateTable(int rowCount, int columnCount)
		{
			if (rowCount < 1 || columnCount < 1)
				throw new ArgumentOutOfRangeException("Row and Column count must be greater than zero.");

			return new Table(this, HelperFunctions.CreateTable(rowCount, columnCount)) { packagePart = packagePart };
		}

		/// <summary>
		/// Releases all resources used by this document.
		/// </summary>
		public void Dispose()
		{
			(package as IDisposable)?.Dispose();

			package = null;
		}

		/// <summary>
		/// Returns the type of editing protection imposed on this document.
		/// </summary>
		/// <returns>The type of editing protection imposed on this document.</returns>
		public EditRestrictions GetProtectionType()
		{
			if (IsProtected)
			{
				XElement documentProtection = settings.Descendants(DocxNamespace.Main + Text_DocumentProtection).FirstOrDefault();
				string editValue = documentProtection.Attribute(DocxNamespace.Main + "edit").Value;
				return Enum.Parse<EditRestrictions>(editValue, ignoreCase: true);
			}

			return EditRestrictions.None;
		}

		public List<Section> GetSections()
		{
			var allParas = Paragraphs;

			var parasInASection = new List<Paragraph>();
			var sections = new List<Section>();

			foreach (var para in allParas)
			{
				var sectionInPara = para.Xml.Descendants().FirstOrDefault(s => s.Name.LocalName == Text_SectionProperties);

				if (sectionInPara == null)
				{
					parasInASection.Add(para);
				}
				else
				{
					parasInASection.Add(para);
					var section = new Section(Document, sectionInPara) { SectionParagraphs = parasInASection };
					sections.Add(section);
					parasInASection = new List<Paragraph>();
				}
			}

			var baseSection = new Section(Document, SectPr) { SectionParagraphs = parasInASection };
			sections.Add(baseSection);

			return sections;
		}

		/// <summary>
		/// Insert a chart in document
		/// </summary>
		public void InsertChart(Chart chart)
		{
			int chartIndex = 1;

			// Create a new chart part uri.
			string chartPartUriPath;

			do
			{
				chartPartUriPath = $"/word/charts/chart{chartIndex}.xml";
				chartIndex++;
			} while (package.PartExists(new Uri(chartPartUriPath, UriKind.Relative)));

			// Create chart part.
			PackagePart chartPackagePart = package.CreatePart(new Uri(chartPartUriPath, UriKind.Relative), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", CompressionOption.Normal);

			// Create a new chart relationship
			string relID = GetNextFreeRelationshipID();
			_ = packagePart.CreateRelationship(chartPackagePart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", relID);

			// Save a chart info the chartPackagePart
			using (TextWriter tw = new StreamWriter(chartPackagePart.GetStream(FileMode.Create, FileAccess.Write)))
				chart.Xml.Save(tw);

			// Insert a new chart into a paragraph.
			Paragraph p = InsertParagraph();
			XElement chartElement = new XElement(DocxNamespace.Main + "r",
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
									new XAttribute(DocxNamespace.RelatedDoc + "id", relID)
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
		/// </example>
		public void InsertDocument(DocX otherDoc, bool append = true)
		{
			// We don't want to effect the origional XDocument, so create a new one from the old one.
			XDocument remote_mainDoc = new XDocument(otherDoc.mainDoc);

			XDocument remote_footnotes = null;
			if (otherDoc.footnotes != null)
				remote_footnotes = new XDocument(otherDoc.footnotes);

			XDocument remote_endnotes = null;
			if (otherDoc.endnotes != null)
				remote_endnotes = new XDocument(otherDoc.endnotes);

			// Remove all header and footer references.
			remote_mainDoc.Descendants(DocxNamespace.Main + "headerReference").Remove();
			remote_mainDoc.Descendants(DocxNamespace.Main + "footerReference").Remove();

			// Get the body of the remote document.
			XElement remote_body = remote_mainDoc.Root.Element(DocxNamespace.Main + Text_Body);

			// Every file that is missing from the local document will have to be copied, every file that already exists will have to be merged.
			PackagePartCollection ppc = otherDoc.package.GetParts();

			List<string> ignoreContentTypes = new List<string>
			{
				"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
				"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
				"application/vnd.openxmlformats-package.core-properties+xml",
				"application/vnd.openxmlformats-officedocument.extended-properties+xml",
				"application/vnd.openxmlformats-package.relationships+xml",
			};

			List<string> imageContentTypes = new List<string>
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
			foreach (PackagePart remote_pp in ppc)
			{
				if (ignoreContentTypes.Contains(remote_pp.ContentType) || imageContentTypes.Contains(remote_pp.ContentType))
					continue;

				// If this external PackagePart already exits then we must merge them.
				if (package.PartExists(remote_pp.Uri))
				{
					PackagePart local_pp = package.GetPart(remote_pp.Uri);
					switch (remote_pp.ContentType)
					{
						case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
							MergeCustoms(remote_pp, local_pp);
							break;

						// Merge footnotes (and endnotes) before merging styles, then set the remote_footnotes to the just updated footnotes
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
							MergeFootnotes(remote_mainDoc, remote_footnotes);
							remote_footnotes = footnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
							MergeEndnotes(remote_mainDoc, remote_endnotes);
							remote_endnotes = endnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
						case "application/vnd.ms-word.stylesWithEffects+xml":
							MergeStyles(remote_mainDoc, otherDoc, remote_footnotes, remote_endnotes);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
							MergeFonts(otherDoc);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
							MergeNumbering(remote_mainDoc, otherDoc);
							break;
					}
				}

				// If this external PackagePart does not exits in the internal document then we can simply copy it.
				else
				{
					var packagePart = ClonePackagePart(remote_pp);
					switch (remote_pp.ContentType)
					{
						case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
							endnotesPart = packagePart;
							endnotes = remote_endnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
							footnotesPart = packagePart;
							footnotes = remote_footnotes;
							break;

						case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
							stylesPart = packagePart;
							using (TextReader tr = new StreamReader(stylesPart.GetStream()))
								styles = XDocument.Load(tr);
							break;

						case "application/vnd.ms-word.stylesWithEffects+xml":
							stylesWithEffectsPart = packagePart;
							using (TextReader tr = new StreamReader(stylesWithEffectsPart.GetStream()))
								stylesWithEffects = XDocument.Load(tr);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
							fontTablePart = packagePart;
							using (TextReader tr = new StreamReader(fontTablePart.GetStream()))
								fontTable = XDocument.Load(tr);
							break;

						case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
							numberingPart = packagePart;
							using (TextReader tr = new StreamReader(numberingPart.GetStream()))
								numbering = XDocument.Load(tr);
							break;
					}

					ClonePackageRelationship(otherDoc, remote_pp, remote_mainDoc);
				}
			}

			foreach (var hyperlink_rel in otherDoc.packagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"))
			{
				var old_rel_Id = hyperlink_rel.Id;
				var new_rel_Id = packagePart.CreateRelationship(hyperlink_rel.TargetUri, hyperlink_rel.TargetMode, hyperlink_rel.RelationshipType).Id;
				foreach (var hyperlink_ref in remote_mainDoc.Descendants(DocxNamespace.Main + "hyperlink"))
				{
					XAttribute a0 = hyperlink_ref.Attribute(DocxNamespace.RelatedDoc + "id");
					if (a0 != null && a0.Value == old_rel_Id)
					{
						a0.SetValue(new_rel_Id);
					}
				}
			}

			////ole object links
			foreach (var oleObject_rel in otherDoc.packagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"))
			{
				var old_rel_Id = oleObject_rel.Id;
				var new_rel_Id = packagePart.CreateRelationship(oleObject_rel.TargetUri, oleObject_rel.TargetMode, oleObject_rel.RelationshipType).Id;
				foreach (var oleObject_ref in remote_mainDoc.Descendants(XName.Get("OLEObject", "urn:schemas-microsoft-com:office:office")))
				{
					XAttribute a0 = oleObject_ref.Attribute(DocxNamespace.RelatedDoc + "id");
					if (a0 != null && a0.Value == old_rel_Id)
					{
						a0.SetValue(new_rel_Id);
					}
				}
			}

			foreach (PackagePart remote_pp in ppc)
			{
				if (imageContentTypes.Contains(remote_pp.ContentType))
				{
					MergeImages(remote_pp, otherDoc, remote_mainDoc, remote_pp.ContentType);
				}
			}

			int id = 0;
			foreach (var local_docPr in mainDoc.Root.Descendants(DocxNamespace.WordProcessingDrawing + "docPr"))
			{
				XAttribute a_id = local_docPr.Attribute("id");
				if (a_id != null && int.TryParse(a_id.Value, out int a_id_value) && a_id_value > id)
					id = a_id_value;
			}
			id++;

			// docPr must be sequential
			foreach (var docPr in remote_body.Descendants(DocxNamespace.WordProcessingDrawing + "docPr"))
			{
				docPr.SetAttributeValue("id", id);
				id++;
			}

			// Add the remote documents contents to this document.
			XElement local_body = mainDoc.Root.Element(DocxNamespace.Main + Text_Body);
			if (append)
				local_body.Add(remote_body.Elements());
			else
				local_body.AddFirst(remote_body.Elements());

			// Copy any missing root attributes to the local document.
			foreach (XAttribute a in remote_mainDoc.Root.Attributes())
			{
				if (mainDoc.Root.Attribute(a.Name) == null)
				{
					mainDoc.Root.SetAttributeValue(a.Name, a.Value);
				}
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
			var toc = TableOfContents.CreateTableOfContents(this, title, switches, headerStyle, maxIncludeLevel, rightTabPos);
			reference.Xml.AddBeforeSelf(toc.Xml);
			return toc;
		}

		/// <summary>
		/// Remove editing protection from this document.
		/// </summary>
		public void RemoveProtection() => settings
				.Descendants(DocxNamespace.Main + Text_DocumentProtection)
				.Remove();

		/// <summary>
		/// Save this document back to the location it was loaded from.
		/// </summary>
		public void Save()
		{
			var xdecl = new XDeclaration("1.0", "UTF-8", "yes");

			// Save the main document
			using (TextWriter tw = new StreamWriter(packagePart.GetStream(FileMode.Create, FileAccess.Write)))
				mainDoc.Save(tw, SaveOptions.None);

			// Create a settings object if necessary.
			if (settings == null)
			{
				using TextReader tr = new StreamReader(settingsPart.GetStream());
				settings = XDocument.Load(tr);
			}

			var sectPr = mainDoc.Root.Element(DocxNamespace.Main + Text_Body).Descendants(DocxNamespace.Main + Text_SectionProperties).FirstOrDefault();
			if (sectPr != null)
			{
				var evenHeaderRef =
				(
					from e in mainDoc.Descendants(DocxNamespace.Main + "headerReference")
					let type = e.Attribute(DocxNamespace.Main + "type")
					where type?.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase) == true
					select e.Attribute(DocxNamespace.RelatedDoc + "id").Value
				 ).LastOrDefault();

				if (evenHeaderRef != null)
				{
					Uri target = PackUriHelper.ResolvePartUri(packagePart.Uri, packagePart.GetRelationship(evenHeaderRef).TargetUri);
					using TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write));
					new XDocument(xdecl, Headers.Even.Xml).Save(tw, SaveOptions.None);
				}

				var oddHeaderRef =
				(
					from e in mainDoc.Descendants(DocxNamespace.Main + "headerReference")
					let type = e.Attribute(DocxNamespace.Main + "type")
					where type?.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase) == true
					select e.Attribute(DocxNamespace.RelatedDoc + "id").Value
				 ).LastOrDefault();

				if (oddHeaderRef != null)
				{
					Uri target = PackUriHelper.ResolvePartUri(packagePart.Uri, packagePart.GetRelationship(oddHeaderRef).TargetUri);
					using TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write));
					new XDocument(xdecl, Headers.Odd.Xml).Save(tw, SaveOptions.None);
				}

				var firstHeaderRef =
				(
					from e in mainDoc.Descendants(DocxNamespace.Main + "headerReference")
					let type = e.Attribute(DocxNamespace.Main + "type")
					where type?.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase) == true
					select e.Attribute(DocxNamespace.RelatedDoc + "id").Value
				 ).LastOrDefault();

				if (firstHeaderRef != null)
				{
					Uri target = PackUriHelper.ResolvePartUri(packagePart.Uri, packagePart.GetRelationship(firstHeaderRef).TargetUri);
					using TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write));
					new XDocument(xdecl, Headers.First.Xml).Save(tw, SaveOptions.None);
				}

				var oddFooterRef =
				(
					from e in mainDoc.Descendants(DocxNamespace.Main + "footerReference")
					let type = e.Attribute(DocxNamespace.Main + "type")
					where type?.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase) == true
					select e.Attribute(DocxNamespace.RelatedDoc + "id").Value
				 ).LastOrDefault();

				if (oddFooterRef != null)
				{
					Uri target = PackUriHelper.ResolvePartUri(packagePart.Uri, packagePart.GetRelationship(oddFooterRef).TargetUri);
					using TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write));
					new XDocument(xdecl, Footers.Odd.Xml).Save(tw, SaveOptions.None);
				}

				var evenFooterRef =
				(
					from e in mainDoc.Descendants(DocxNamespace.Main + "footerReference")
					let type = e.Attribute(DocxNamespace.Main + "type")
					where type?.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase) == true
					select e.Attribute(DocxNamespace.RelatedDoc + "id").Value
				 ).LastOrDefault();

				if (evenFooterRef != null)
				{
					Uri target = PackUriHelper.ResolvePartUri(packagePart.Uri, packagePart.GetRelationship(evenFooterRef).TargetUri);
					using TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write));
					new XDocument(xdecl, Footers.Even.Xml).Save(tw, SaveOptions.None);
				}

				var firstFooterRef =
				(
					 from e in mainDoc.Descendants(DocxNamespace.Main + "footerReference")
					 let type = e.Attribute(DocxNamespace.Main + "type")
					 where type?.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase) == true
					 select e.Attribute(DocxNamespace.RelatedDoc + "id").Value
				).LastOrDefault();

				if (firstFooterRef != null)
				{
					Uri target = PackUriHelper.ResolvePartUri(packagePart.Uri, packagePart.GetRelationship(firstFooterRef).TargetUri);
					using TextWriter tw = new StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write));
					new XDocument(xdecl, Footers.First.Xml).Save(tw, SaveOptions.None);
				}

				// Save the settings document.
				using (TextWriter tw = new StreamWriter(settingsPart.GetStream(FileMode.Create, FileAccess.Write)))
					settings.Save(tw, SaveOptions.None);

				if (endnotesPart != null)
				{
					using TextWriter tw = new StreamWriter(endnotesPart.GetStream(FileMode.Create, FileAccess.Write));
					endnotes.Save(tw, SaveOptions.None);
				}

				if (footnotesPart != null)
				{
					using TextWriter tw = new StreamWriter(footnotesPart.GetStream(FileMode.Create, FileAccess.Write));
					footnotes.Save(tw, SaveOptions.None);
				}

				if (stylesPart != null)
				{
					using TextWriter tw = new StreamWriter(stylesPart.GetStream(FileMode.Create, FileAccess.Write));
					styles.Save(tw, SaveOptions.None);
				}

				if (stylesWithEffectsPart != null)
				{
					using TextWriter tw = new StreamWriter(stylesWithEffectsPart.GetStream(FileMode.Create, FileAccess.Write));
					stylesWithEffects.Save(tw, SaveOptions.None);
				}

				if (numberingPart != null)
				{
					using TextWriter tw = new StreamWriter(numberingPart.GetStream(FileMode.Create, FileAccess.Write));
					numbering.Save(tw, SaveOptions.None);
				}

				if (fontTablePart != null)
				{
					using TextWriter tw = new StreamWriter(fontTablePart.GetStream(FileMode.Create, FileAccess.Write));
					fontTable.Save(tw, SaveOptions.None);
				}
			}

			// Close the package and commit changes to the memory stream.
			package.Close();

			// Save back to the file or stream
			if (filename != null)
			{
				memoryStream.Seek(0, SeekOrigin.Begin);
				using FileStream fs = File.Create(filename);
				memoryStream.WriteTo(fs);
			}
			else
			{
				if (stream.CanSeek)
				{
					stream.SetLength(0);
					stream.Seek(0, SeekOrigin.Begin);
				}
				memoryStream.WriteTo(stream);
			}

			// Reopen the package.
			package = Package.Open(memoryStream, FileMode.Open, FileAccess.ReadWrite);
			LoadDocumentParts();
		}

		/// <summary>
		/// Save this document to a file.
		/// </summary>
		/// <param name="filename">The filename to save this document as.</param>
		public void SaveAs(string filename)
		{
			this.filename = filename;
			this.stream = null;
			Save();
		}

		/// <summary>
		/// Save this document to a Stream.
		/// </summary>
		/// <param name="stream">The Stream to save this document to.</param>
		public void SaveAs(Stream stream)
		{
			this.filename = null;
			this.stream = stream;
			Save();
		}

		internal static void PostCreation(Package package, DocumentTypes documentType = DocumentTypes.Document)
		{
			// Create the main document part for this package
			PackagePart mainDocumentPart = package.CreatePart(new Uri("/word/document.xml", UriKind.Relative),
				documentType == DocumentTypes.Document ? DocxContentType.Document : DocxContentType.Template,
				CompressionOption.Normal);
			package.CreateRelationship(mainDocumentPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");

			// Load the document part into a XDocument object
			XDocument mainDoc;
			using (TextReader tr = new StreamReader(mainDocumentPart.GetStream(FileMode.Create, FileAccess.ReadWrite)))
			{
				mainDoc = XDocument.Parse
				(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
				   <w:document xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
				   <w:body>
					<w:sectPr w:rsidR=""003E25F4"" w:rsidSect=""00FC3028"">
						<w:pgSz w:w=""11906"" w:h=""16838""/>
						<w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""708"" w:footer=""708"" w:gutter=""0""/>
						<w:cols w:space=""708""/>
						<w:docGrid w:linePitch=""360""/>
					</w:sectPr>
				   </w:body>
				   </w:document>"
				);
			}

			// Save the main document
			using (TextWriter tw = new StreamWriter(mainDocumentPart.GetStream(FileMode.Create, FileAccess.Write)))
				mainDoc.Save(tw, SaveOptions.None);

			_ = HelperFunctions.AddDefaultStylesXml(package);
			_ = HelperFunctions.AddDefaultNumberingXml(package);

			package.Close();
		}

		internal static DocX PostLoad(ref Package package)
		{
			DocX document = new DocX { package = package };
			document.Document = document;
			document.LoadDocumentParts();
			return document;
		}

		internal static void UpdateCorePropertyValue(DocX document, string corePropertyName, string corePropertyValue)
		{
			string matchPattern = string.Format(@"(DOCPROPERTY)?{0}\\\*MERGEFORMAT", corePropertyName).ToLower();
			foreach (XElement e in document.mainDoc.Descendants(DocxNamespace.Main + "fldSimple"))
			{
				string attr_value = e.Attribute(DocxNamespace.Main + "instr").Value.Replace(" ", string.Empty).Trim().ToLower();

				if (Regex.IsMatch(attr_value, matchPattern))
				{
					XElement firstRun = e.Element(DocxNamespace.Main + "r");
					XElement firstText = firstRun.Element(DocxNamespace.Main + "t");
					XElement rPr = firstText.Element(DocxNamespace.Main + "rPr");

					// Delete everything and insert updated text value
					e.RemoveNodes();

					XElement t = new XElement(DocxNamespace.Main + "t", rPr, corePropertyValue);
					TextBlock.PreserveSpace(t);
					e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(DocxNamespace.Main + "rPr"), t));
				}
			}

			IEnumerable<PackagePart> headerParts = from headerPart in document.package.GetParts()
												   where (Regex.IsMatch(headerPart.Uri.ToString(), @"/word/header\d?.xml"))
												   select headerPart;
			foreach (PackagePart pp in headerParts)
			{
				XDocument header = XDocument.Load(new StreamReader(pp.GetStream()));

				foreach (XElement e in header.Descendants(DocxNamespace.Main + "fldSimple"))
				{
					string attr_value = e.Attribute(DocxNamespace.Main + "instr").Value.Replace(" ", string.Empty).Trim().ToLower();
					if (Regex.IsMatch(attr_value, matchPattern))
					{
						XElement firstRun = e.Element(DocxNamespace.Main + "r");

						// Delete everything and insert updated text value
						e.RemoveNodes();

						XElement t = new XElement(DocxNamespace.Main + "t", corePropertyValue);
						TextBlock.PreserveSpace(t);
						e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(DocxNamespace.Main + "rPr"), t));
					}
				}

				using TextWriter tw = new StreamWriter(pp.GetStream(FileMode.Create, FileAccess.Write));
				header.Save(tw);
			}

			IEnumerable<PackagePart> footerParts = from footerPart in document.package.GetParts()
												   where (Regex.IsMatch(footerPart.Uri.ToString(), @"/word/footer\d?.xml"))
												   select footerPart;
			foreach (PackagePart pp in footerParts)
			{
				XDocument footer = XDocument.Load(new StreamReader(pp.GetStream()));

				foreach (XElement e in footer.Descendants(DocxNamespace.Main + "fldSimple"))
				{
					string attr_value = e.Attribute(DocxNamespace.Main + "instr").Value.Replace(" ", string.Empty).Trim().ToLower();
					if (Regex.IsMatch(attr_value, matchPattern))
					{
						XElement firstRun = e.Element(DocxNamespace.Main + "r");

						// Delete everything and insert updated text value
						e.RemoveNodes();

						XElement t = new XElement(DocxNamespace.Main + "t", corePropertyValue);
						TextBlock.PreserveSpace(t);
						e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(DocxNamespace.Main + "rPr"), t));
					}
				}

				using TextWriter tw = new StreamWriter(pp.GetStream(FileMode.Create, FileAccess.Write));
				footer.Save(tw);
			}

			PopulateDocument(document, document.package);
		}

		/// <summary>
		/// Update the custom properties inside the document
		/// </summary>
		/// <param name="document">The DocX document</param>
		/// <param name="customPropertyName">The property used inside the document</param>
		/// <param name="customPropertyValue">The new value for the property</param>
		/// <remarks>Different version of Word create different Document XML.</remarks>
		internal static void UpdateCustomPropertyValue(DocX document, string customPropertyName, string customPropertyValue)
		{
			// A list of documents, which will contain, The Main Document and if they exist: header1, header2, header3, footer1, footer2, footer3.
			List<XElement> documents = new List<XElement> { document.mainDoc.Root };

			// Check if each header exists and add if if so.

			Headers headers = document.Headers;
			if (headers.First != null)
				documents.Add(headers.First.Xml);
			if (headers.Odd != null)
				documents.Add(headers.Odd.Xml);
			if (headers.Even != null)
				documents.Add(headers.Even.Xml);

			// Check if each footer exists and add if if so.

			Footers footers = document.Footers;
			if (footers.First != null)
				documents.Add(footers.First.Xml);
			if (footers.Odd != null)
				documents.Add(footers.Odd.Xml);
			if (footers.Even != null)
				documents.Add(footers.Even.Xml);

			var matchCustomPropertyName = customPropertyName;
			if (customPropertyName.Contains(" ")) matchCustomPropertyName = "\"" + customPropertyName + "\"";
			string match_value = string.Format(@"DOCPROPERTY  {0}  \* MERGEFORMAT", matchCustomPropertyName).Replace(" ", string.Empty);

			// Process each document in the list.
			foreach (XElement doc in documents)
			{
				foreach (XElement e in doc.Descendants(DocxNamespace.Main + "instrText"))
				{
					string attr_value = e.Value.Replace(" ", string.Empty).Trim();

					if (attr_value.Equals(match_value, StringComparison.CurrentCultureIgnoreCase))
					{
						XNode node = e.Parent.NextNode;
						bool found = false;
						while (true)
						{
							if (node.NodeType == XmlNodeType.Element)
							{
								var ele = node as XElement;
								var match = ele.Descendants(DocxNamespace.Main + "t");
								if (match.Any())
								{
									if (!found)
									{
										match.First().Value = customPropertyValue;
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
										{
											break;
										}
									}
								}
							}
							node = node.NextNode;
						}
					}
				}

				foreach (XElement e in doc.Descendants(DocxNamespace.Main + "fldSimple"))
				{
					string attr_value = e.Attribute(DocxNamespace.Main + "instr").Value.Replace(" ", string.Empty).Trim();

					if (attr_value.Equals(match_value, StringComparison.CurrentCultureIgnoreCase))
					{
						XElement firstRun = e.Element(DocxNamespace.Main + "r");
						XElement firstText = firstRun.Element(DocxNamespace.Main + "t");
						XElement rPr = firstText.Element(DocxNamespace.Main + "rPr");

						// Delete everything and insert updated text value
						e.RemoveNodes();

						XElement t = new XElement(DocxNamespace.Main + "t", rPr, customPropertyValue);
						TextBlock.PreserveSpace(t);
						e.Add(new XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(DocxNamespace.Main + "rPr"), t));
					}
				}
			}
		}

		/// <summary>
		/// Adds a Header to a document.
		/// If the document already contains a Header it will be replaced.
		/// </summary>
		/// <returns>The Header that was added to the document.</returns>
		internal void AddHeadersOrFooters(bool b)
		{
			string element = "ftr";
			string reference = "footer";
			if (b)
			{
				element = "hdr";
				reference = "header";
			}

			DeleteHeadersOrFooters(b);

			for (int i = 1; i < 4; i++)
			{
				string header_uri = string.Format("/word/{0}{1}.xml", reference, i);

				PackagePart headerPart = package.CreatePart(new Uri(header_uri, UriKind.Relative), string.Format("application/vnd.openxmlformats-officedocument.wordprocessingml.{0}+xml", reference), CompressionOption.Normal);
				PackageRelationship headerRelationship = packagePart.CreateRelationship(headerPart.Uri, TargetMode.Internal, string.Format("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference));

				XDocument header;

				// Load the document part into a XDocument object
				using (TextReader tr = new StreamReader(headerPart.GetStream(FileMode.Create, FileAccess.ReadWrite)))
				{
					header = XDocument.Parse
					(string.Format(@"<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>
					   <w:{0} xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">
						 <w:p w:rsidR=""009D472B"" w:rsidRDefault=""009D472B"">
						   <w:pPr>
							 <w:pStyle w:val=""{1}"" />
						   </w:pPr>
						 </w:p>
					   </w:{0}>", element, reference)
					);
				}

				// Save the main document
				using (TextWriter tw = new StreamWriter(headerPart.GetStream(FileMode.Create, FileAccess.Write)))
					header.Save(tw, SaveOptions.None);

				string type = i switch
				{
					1 => "default",
					2 => "even",
					3 => "first",
					_ => throw new ArgumentOutOfRangeException(),
				};

				SectPr.Add
				(
					new XElement(DocxNamespace.Main + string.Format("{0}Reference", reference),
						new XAttribute(DocxNamespace.Main + "type", type),
						new XAttribute(DocxNamespace.RelatedDoc + "id", headerRelationship.Id)
					)
				);
			}
		}

		internal void AddHyperlinkStyleIfNotPresent()
		{
			// If the document contains no /word/styles.xml create one and associate it
			if (!package.PartExists(DocxSections.StylesUri))
				HelperFunctions.AddDefaultStylesXml(package);

			// Ensure we are looking at the correct one.
			if (stylesPart == null)
			{
				stylesPart = package.GetPart(DocxSections.StylesUri);
			}

			// Load the styles.xml into memory.
			if (styles == null)
			{
				using var sr = new StreamReader(stylesPart.GetStream(FileMode.Open, FileAccess.Read));
				styles = XDocument.Load(sr);
			}

			// Check for the "hyperlinkStyle"
			bool hyperlinkStyleExists = styles.Element(DocxNamespace.Main + "styles").Elements()
				.Any(e => e.Attribute(DocxNamespace.Main + "styleId")?.Value == "Hyperlink");

			if (!hyperlinkStyleExists)
			{
				// Add a simple Hyperlink style (blue + underline + default font + size)
				styles.Element(DocxNamespace.Main + "styles").Add(
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

		internal Image AddImage(object imageData, string contentType = "image/jpeg")
		{
			// Open a Stream to the new image being added.
			Stream newImageStream;
			if (imageData is string)
				newImageStream = new FileStream(imageData as string, FileMode.Open, FileAccess.Read);
			else
				newImageStream = imageData as Stream;

			// Get all image parts in word\document.xml

			PackageRelationshipCollection relationshipsByImages = packagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
			List<PackagePart> imageParts = relationshipsByImages.Select(ir => package.GetParts()
						.FirstOrDefault(p => p.Uri.ToString().EndsWith(ir.TargetUri.ToString())))
						.Where(e => e != null)
						.ToList();

			foreach (PackagePart relsPart in package.GetParts().Where(part => part.Uri.ToString().Contains("/word/") 
						&& part.ContentType.Equals("application/vnd.openxmlformats-package.relationships+xml")))
			{
				XDocument relsPartContent;
				using (TextReader tr = new StreamReader(relsPart.GetStream(FileMode.Open, FileAccess.Read)))
					relsPartContent = XDocument.Load(tr);

				var imageRelationships = relsPartContent.Root.Elements()
											.Where(imageRel => imageRel.Attribute("Type")
													.Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"));

				foreach (XElement imageRelationship in imageRelationships)
				{
					if (imageRelationship.Attribute("Target") != null)
					{
						string targetMode = imageRelationship.AttributeValue("TargetMode");
						if (!targetMode.Equals("External"))
						{
							string imagePartUri = Path.Combine(Path.GetDirectoryName(relsPart.Uri.ToString()), imageRelationship.AttributeValue("Target"));
							imagePartUri = Path.GetFullPath(imagePartUri.Replace("\\_rels", string.Empty));
							imagePartUri = imagePartUri.Replace(Path.GetFullPath("\\"), string.Empty).Replace("\\", "/");

							if (!imagePartUri.StartsWith("/"))
								imagePartUri = "/" + imagePartUri;

							PackagePart imagePart = package.GetPart(new Uri(imagePartUri, UriKind.Relative));
							imageParts.Add(imagePart);
						}
					}
				}
			}

			// Loop through each image part in this document.
			foreach (PackagePart pp in imageParts)
			{
				// Open a tempory Stream to this image part.
				using (Stream tempStream = pp.GetStream(FileMode.Open, FileAccess.Read))
				{
					// Compare this image to the new image being added.
					if (HelperFunctions.IsSameFile(tempStream, newImageStream))
					{
						// Get the image object for this image part
						string id = packagePart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
											   .Where(r => r.TargetUri == pp.Uri)
											   .Select(r => r.Id).First();

						// Return the Image object
						return Images.First(i => i.Id == id);
					}
				}
			}

			string imgPartUriPath = string.Empty;
			string extension = contentType.Substring(contentType.LastIndexOf("/") + 1);

			// Get a unique imgPartUriPath
			do
			{
				imgPartUriPath = $"/word/media/{Guid.NewGuid()}.{extension}";
			} while (package.PartExists(new Uri(imgPartUriPath, UriKind.Relative)));

			PackagePart img = package.CreatePart(new Uri(imgPartUriPath, UriKind.Relative), contentType, CompressionOption.Normal);

			// Create a new image relationship
			PackageRelationship rel = packagePart.CreateRelationship(img.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

			// Open a Stream to the newly created Image part.
			using (Stream stream = img.GetStream(FileMode.Create, FileAccess.Write))
			{
				// Using the Stream to the real image, copy this streams data into the newly create Image part.
				using (newImageStream)
				{
					if (newImageStream.CanSeek)
					{
						newImageStream.Seek(0, SeekOrigin.Begin);
						newImageStream.CopyTo(stream);
					}
					else
					{
						byte[] bytes = new byte[newImageStream.Length];
						newImageStream.Read(bytes, 0, (int)newImageStream.Length);
						stream.Write(bytes, 0, (int)newImageStream.Length);
					}
				}
			}

			return new Image(this, rel);
		}

		internal XDocument AddStylesForList()
		{
			var wordStylesUri = new Uri("/word/styles.xml", UriKind.Relative);

			// If the internal document contains no /word/styles.xml create one.
			if (!package.PartExists(wordStylesUri))
				HelperFunctions.AddDefaultStylesXml(package);

			// Load the styles.xml into memory.
			XDocument wordStyles;
			using (TextReader tr = new StreamReader(package.GetPart(wordStylesUri).GetStream()))
				wordStyles = XDocument.Load(tr);

			bool listStyleExists =
			(
			  from s in wordStyles.Element(DocxNamespace.Main + "styles").Elements()
			  let styleId = s.Attribute(DocxNamespace.Main + "styleId")
			  where (styleId?.Value == "ListParagraph")
			  select s
			).Any();

			if (!listStyleExists)
			{
				var style = new XElement
				(
					DocxNamespace.Main + "style",
					new XAttribute(DocxNamespace.Main + "type", "paragraph"),
					new XAttribute(DocxNamespace.Main + "styleId", "ListParagraph"),
						new XElement(DocxNamespace.Main + "name", new XAttribute(DocxNamespace.Main + "val", "List Paragraph")),
						new XElement(DocxNamespace.Main + "basedOn", new XAttribute(DocxNamespace.Main + "val", "Normal")),
						new XElement(DocxNamespace.Main + "uiPriority", new XAttribute(DocxNamespace.Main + "val", "34")),
						new XElement(DocxNamespace.Main + "qformat"),
						new XElement(DocxNamespace.Main + "rsid", new XAttribute(DocxNamespace.Main + "val", "00832EE1")),
						new XElement
						(
							DocxNamespace.Main + "rPr",
							new XElement(DocxNamespace.Main + "ind", new XAttribute(DocxNamespace.Main + Text_Left, "720")),
							new XElement
							(
								DocxNamespace.Main + "contextualSpacing"
							)
						)
				);
				wordStyles.Element(DocxNamespace.Main + "styles").Add(style);

				// Save the styles document.
				using TextWriter tw = new StreamWriter(package.GetPart(wordStylesUri).GetStream());
				wordStyles.Save(tw);
			}

			return wordStyles;
		}

		internal void DeleteHeadersOrFooters(bool b)
		{
			string reference = "footer";
			if (b)
				reference = "header";

			// Get all header Relationships in this document.
			foreach (PackageRelationship header_relationship in packagePart.GetRelationshipsByType(string.Format("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference)))
			{
				// Get the TargetUri for this Part.
				Uri header_uri = header_relationship.TargetUri;

				// Check to see if the document actually contains the Part.
				if (!header_uri.OriginalString.StartsWith("/word/"))
					header_uri = new Uri("/word/" + header_uri.OriginalString, UriKind.Relative);

				if (package.PartExists(header_uri))
				{
					// Delete the Part
					package.DeletePart(header_uri);

					// Get all references to this Relationship in the document.
					var query =
					(
						from e in mainDoc.Descendants(DocxNamespace.Main + Text_Body).Descendants()
						where (e.Name.LocalName == string.Format("{0}Reference", reference)) && (e.Attribute(DocxNamespace.RelatedDoc + "id").Value == header_relationship.Id)
						select e
					);

					// Remove all references to this Relationship in the document.
					for (int i = 0; i < query.Count(); i++)
						query.ElementAt(i).Remove();

					// Delete the Relationship.
					package.DeleteRelationship(header_relationship.Id);
				}
			}
		}

		internal string GetCollectiveText(List<PackagePart> list)
		{
			string text = string.Empty;

			foreach (var hp in list)
			{
				using TextReader tr = new StreamReader(hp.GetStream());
				XDocument d = XDocument.Load(tr);

				StringBuilder sb = new StringBuilder();

				// Loop through each text item in this run
				foreach (XElement descendant in d.Descendants())
				{
					switch (descendant.Name.LocalName)
					{
						case "tab":
							sb.Append("\t");
							break;

						case "br":
							sb.Append("\n");
							break;

						case "t":
						case "delText":
							sb.Append(descendant.Value);
							break;
					}
				}

				text += "\n" + sb.ToString();
			}

			return text;
		}

		internal void LoadDocumentParts()
		{
			packagePart = package.GetParts().Single(p =>
				p.ContentType.Equals(DocxContentType.Document, StringComparison.CurrentCultureIgnoreCase) ||
				p.ContentType.Equals(DocxContentType.Template, StringComparison.CurrentCultureIgnoreCase));

			using (TextReader tr = new StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read)))
				mainDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace);

			PopulateDocument(this, package);

			using (TextReader tr = new StreamReader(settingsPart.GetStream()))
				settings = XDocument.Load(tr);

			paragraphLookup.Clear();
			foreach (var paragraph in Paragraphs)
			{
				if (!paragraphLookup.ContainsKey(paragraph.endIndex))
					paragraphLookup.Add(paragraph.endIndex, paragraph);
			}
		}

		private static void PopulateDocument(DocX document, Package package)
		{
			Headers headers = new Headers
			{
				Odd = document.GetHeaderByType("default"),
				Even = document.GetHeaderByType("even"),
				First = document.GetHeaderByType("first")
			};

			Footers footers = new Footers
			{
				Odd = document.GetFooterByType("default"),
				Even = document.GetFooterByType("even"),
				First = document.GetFooterByType("first")
			};

			document.Xml = document.mainDoc.Root.Element(DocxNamespace.Main + Text_Body);
			document.Headers = headers;
			document.Footers = footers;
			document.settingsPart = HelperFunctions.CreateOrGetSettingsPart(package);

			_ = package.GetParts();

			foreach (var rel in document.packagePart.GetRelationships())
			{
				switch (rel.RelationshipType)
				{
					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes":
						document.endnotesPart = package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						using (TextReader tr = new StreamReader(document.endnotesPart.GetStream()))
							document.endnotes = XDocument.Load(tr);
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes":
						document.footnotesPart = package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						using (TextReader tr = new StreamReader(document.footnotesPart.GetStream()))
							document.footnotes = XDocument.Load(tr);
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
						document.stylesPart = package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						using (TextReader tr = new StreamReader(document.stylesPart.GetStream()))
							document.styles = XDocument.Load(tr);
						break;

					case "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects":
						document.stylesWithEffectsPart = package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						using (TextReader tr = new StreamReader(document.stylesWithEffectsPart.GetStream()))
							document.stylesWithEffects = XDocument.Load(tr);
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
						document.fontTablePart = package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						using (TextReader tr = new StreamReader(document.fontTablePart.GetStream()))
							document.fontTable = XDocument.Load(tr);
						break;

					case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering":
						document.numberingPart = package.GetPart(new Uri("/word/" + rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative));
						using (TextReader tr = new StreamReader(document.numberingPart.GetStream()))
							document.numbering = XDocument.Load(tr);
						break;

					default:
						break;
				}
			}
		}

		private PackagePart ClonePackagePart(PackagePart pp)
		{
			PackagePart new_pp = package.CreatePart(pp.Uri, pp.ContentType, CompressionOption.Normal);

			using (Stream s_read = pp.GetStream())
			{
				using Stream s_write = new_pp.GetStream(FileMode.Create);
				byte[] buffer = new byte[short.MaxValue];
				int read;
				while ((read = s_read.Read(buffer, 0, buffer.Length)) > 0)
				{
					s_write.Write(buffer, 0, read);
				}
			}

			return new_pp;
		}

		private void ClonePackageRelationship(DocX remote_document, PackagePart pp, XDocument remote_mainDoc)
		{
			string url = pp.Uri.OriginalString.Replace("/", "");
			foreach (var remote_rel in remote_document.packagePart.GetRelationships())
			{
				if (url.Equals("word" + remote_rel.TargetUri.OriginalString.Replace("/", "")))
				{
					string remote_Id = remote_rel.Id;
					string local_Id = packagePart.CreateRelationship(remote_rel.TargetUri, remote_rel.TargetMode, remote_rel.RelationshipType).Id;

					// Replace all instances of remote_Id in the local document with local_Id
					foreach (var elem in remote_mainDoc.Descendants(DocxNamespace.DrawingMain + "blip"))
					{
						XAttribute embed = elem.Attribute(DocxNamespace.RelatedDoc + "embed");
						if (embed?.Value == remote_Id)
						{
							embed.SetValue(local_Id);
						}
					}

					// Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
					foreach (var elem in remote_mainDoc.Descendants(DocxNamespace.VML + "imagedata"))
					{
						XAttribute id = elem.Attribute(DocxNamespace.RelatedDoc + "id");
						if (id?.Value == remote_Id)
						{
							id.SetValue(local_Id);
						}
					}
					break;
				}
			}
		}

		private string ComputeMD5HashString(Stream stream)
		{
			MD5 md5 = MD5.Create();
			byte[] hash = md5.ComputeHash(stream);
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < hash.Length; i++)
				sb.Append(hash[i].ToString("X2"));
			return sb.ToString();
		}

		private Footer GetFooterByType(string type)
		{
			return (Footer)GetHeaderOrFooterByType(type, false);
		}

		private Header GetHeaderByType(string type)
		{
			return (Header)GetHeaderOrFooterByType(type, true);
		}

		private Container GetHeaderOrFooterByType(string type, bool isHeader)
		{
			// Switch which handles either case Header\Footer, this just cuts down on code duplication.
			string reference = (!isHeader) ? "footerReference" : "headerReference";

			// Get the Id of the [default, even or first] [Header or Footer]
			string Id = mainDoc.Descendants(DocxNamespace.Main + Text_Body).Descendants()
				.Where(e => (e.Name.LocalName == reference) && (e.Attribute(DocxNamespace.Main + "type").Value == type))
				.Select(e => e.Attribute(DocxNamespace.RelatedDoc + "id").Value)
				.LastOrDefault();

			if (Id != null)
			{
				// Get the Xml file for this Header or Footer.
				Uri partUri = packagePart.GetRelationship(Id).TargetUri;

				// Weird problem with PackaePart API.
				if (!partUri.OriginalString.StartsWith("/word/"))
					partUri = new Uri("/word/" + partUri.OriginalString, UriKind.Relative);

				// Get the Part and open a stream to get the Xml file.
				PackagePart part = package.GetPart(partUri);

				using TextReader tr = new StreamReader(part.GetStream());
				XDocument doc = XDocument.Load(tr);

				// Header and Footer extend Container.
				return isHeader
					? new Header(this, doc.Element(DocxNamespace.Main + "hdr"), part)
					: (Container)new Footer(this, doc.Element(DocxNamespace.Main + "ftr"), part);
			}

			// If we got this far something went wrong.
			return null;
		}

		private float GetMarginAttribute(XName name)
		{
			XElement pgMar = SectPr.Element(DocxNamespace.Main + Text_PageMargins);
			XAttribute top = pgMar?.Attribute(name);
			if (top != null && float.TryParse(top.Value, out float f))
			{
				return (int)(f / 20.0f);
			}

			return 0;
		}

		private string GetNextFreeRelationshipID()
		{
			int id = (
				 from r in packagePart.GetRelationships()
				 where r.Id.Substring(0, 3).Equals("rId")
				 select int.Parse(r.Id.Substring(3))
			 ).DefaultIfEmpty().Max();

			// The conventiom for ids is rid01, rid02, etc
			string newId = id.ToString();
			if (int.TryParse(newId, out int result))
			{
				return "rId" + (result + 1);
			}
			else
			{
				string guid = string.Empty;
				do
				{
					guid = Guid.NewGuid().ToString();
				} while (char.IsDigit(guid[0]));

				return guid;
			}
		}

		private void MergeImages(PackagePart remote_pp, DocX remote_document, XDocument remote_mainDoc, string contentType)
		{
			// Before doing any other work, check to see if this image is actually referenced in the document.
			var remote_rel = remote_document.packagePart.GetRelationships().Where(r => r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString.Replace("/word/", ""))).FirstOrDefault();
			if (remote_rel == null)
			{
				remote_rel = remote_document.packagePart.GetRelationships().Where(r => r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString)).FirstOrDefault();
				if (remote_rel == null)
					return;
			}

			string remote_Id = remote_rel.Id;
			string remote_hash = ComputeMD5HashString(remote_pp.GetStream());
			var image_parts = package.GetParts().Where(pp => pp.ContentType.Equals(contentType));

			bool found = false;
			foreach (var part in image_parts)
			{
				string local_hash = ComputeMD5HashString(part.GetStream());
				if (local_hash.Equals(remote_hash))
				{
					// This image already exists in this document.
					found = true;

					var local_rel = packagePart.GetRelationships().FirstOrDefault(r => r.TargetUri.OriginalString.Equals(part.Uri.OriginalString.Replace("/word/", "")))
								 ?? packagePart.GetRelationships().FirstOrDefault(r => r.TargetUri.OriginalString.Equals(part.Uri.OriginalString));
					if (local_rel != null)
					{
						string new_Id = local_rel.Id;

						// Replace all instances of remote_Id in the local document with local_Id
						foreach (var elem in remote_mainDoc.Descendants(DocxNamespace.DrawingMain + "blip"))
						{
							XAttribute embed = elem.Attribute(DocxNamespace.RelatedDoc + "embed");
							if (embed != null && embed.Value == remote_Id)
							{
								embed.SetValue(new_Id);
							}
						}

						// Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
						foreach (var elem in remote_mainDoc.Descendants(DocxNamespace.VML + "imagedata"))
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
					new_uri = "/" + new_uri;

				PackagePart new_pp = package.CreatePart(new Uri(new_uri, UriKind.Relative), remote_pp.ContentType, CompressionOption.Normal);

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

				PackageRelationship pr = packagePart.CreateRelationship(new Uri(new_uri, UriKind.Relative), TargetMode.Internal,
													"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

				string new_Id = pr.Id;

				//Check if the remote relationship id is a default rId from Word
				Match defRelId = Regex.Match(remote_Id, @"rId\d+", RegexOptions.IgnoreCase);

				// Replace all instances of remote_Id in the local document with local_Id
				foreach (var elem in remote_mainDoc.Descendants(DocxNamespace.DrawingMain + "blip"))
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
					foreach (var elem in mainDoc.Descendants(DocxNamespace.DrawingMain + "blip"))
					{
						XAttribute embed = elem.Attribute(DocxNamespace.RelatedDoc + "embed");
						if (embed != null && embed.Value == remote_Id)
						{
							embed.SetValue(new_Id);
						}
					}

					// Replace all instances of remote_Id in the local document with local_Id
					foreach (var elem in mainDoc.Descendants(DocxNamespace.VML + "imagedata"))
					{
						XAttribute id = elem.Attribute(DocxNamespace.RelatedDoc + "id");
						if (id != null && id.Value == remote_Id)
						{
							id.SetValue(new_Id);
						}
					}
				}

				// Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
				foreach (var elem in remote_mainDoc.Descendants(DocxNamespace.VML + "imagedata"))
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
			// Get the remote documents custom.xml file.
			XDocument remote_custom_document;
			using (TextReader tr = new StreamReader(remote_pp.GetStream()))
				remote_custom_document = XDocument.Load(tr);

			// Get the local documents custom.xml file.
			XDocument local_custom_document;
			using (TextReader tr = new StreamReader(local_pp.GetStream()))
				local_custom_document = XDocument.Load(tr);

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
						found = true;
				}

				if (!found)
				{
					remote_property.SetAttributeValue("pid", pid);
					local_custom_document.Root.Add(remote_property);
					pid++;
				}
			}

			// Save the modified local custom styles.xml file.
			using TextWriter tw = new StreamWriter(local_pp.GetStream(FileMode.Create, FileAccess.Write));
			local_custom_document.Save(tw, SaveOptions.None);
		}

		private void MergeEndnotes(XDocument remote_mainDoc, XDocument remote_endnotes)
		{
			IEnumerable<int> ids = endnotes.Root.Descendants()
				.Where(d => d.Name.LocalName == "endnote")
				.Select(d => int.Parse(d.Attribute(DocxNamespace.Main + "id").Value));

			int max_id = ids.Max() + 1;
			var endnoteReferences = remote_mainDoc.Descendants(DocxNamespace.Main + "endnoteReference");

			foreach (var endnote in remote_endnotes.Root.Elements().OrderBy(fr => fr.Attribute(DocxNamespace.RelatedDoc + "id")).Reverse())
			{
				XAttribute id = endnote.Attribute(DocxNamespace.Main + "id");
				if (id != null && int.TryParse(id.Value, out int i) && i > 0)
				{
					foreach (var endnoteRef in endnoteReferences)
					{
						XAttribute a = endnoteRef.Attribute(DocxNamespace.Main + "id");
						if (a != null && int.Parse(a.Value).Equals(i))
						{
							a.SetValue(max_id);
						}
					}

					// We care about copying this footnote.
					endnote.SetAttributeValue(DocxNamespace.Main + "id", max_id);
					endnotes.Root.Add(endnote);
					max_id++;
				}
			}
		}

		private void MergeFonts(DocX remote)
		{
			// Add each remote font to this document.
			var remote_fonts = remote.fontTable.Root.Elements(DocxNamespace.Main + "font");
			var local_fonts = fontTable.Root.Elements(DocxNamespace.Main + "font");

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
					fontTable.Root.Add(remote_font);
				}
			}
		}

		private void MergeFootnotes(XDocument remote_mainDoc, XDocument remote_footnotes)
		{
			IEnumerable<int> ids = footnotes.Root.Descendants()
				.Where(d => d.Name.LocalName == "footnote")
				.Select(d => int.Parse(d.Attribute(DocxNamespace.Main + "id").Value));

			int max_id = ids.Max() + 1;
			var footnoteReferences = remote_mainDoc.Descendants(DocxNamespace.Main + "footnoteReference");

			foreach (var footnote in remote_footnotes.Root.Elements().OrderBy(fr => fr.Attribute(DocxNamespace.RelatedDoc + "id")).Reverse())
			{
				XAttribute id = footnote.Attribute(DocxNamespace.Main + "id");
				if (id != null && int.TryParse(id.Value, out int i) && i > 0)
				{
					foreach (var footnoteRef in footnoteReferences)
					{
						XAttribute a = footnoteRef.Attribute(DocxNamespace.Main + "id");
						if (a != null && int.Parse(a.Value).Equals(i))
						{
							a.SetValue(max_id);
						}
					}

					// We care about copying this footnote.
					footnote.SetAttributeValue(DocxNamespace.Main + "id", max_id);
					footnotes.Root.Add(footnote);
					max_id++;
				}
			}
		}

		private void MergeNumbering(XDocument remote_mainDoc, DocX remote)
		{
			// Add each remote numbering to this document.
			var remote_abstractNums = remote.numbering.Root.Elements(DocxNamespace.Main + "abstractNum").ToList();
			int guidd = 0;
			foreach (var an in remote_abstractNums)
			{
				XAttribute a = an.Attribute(DocxNamespace.Main + "abstractNumId");
				if (a != null && int.TryParse(a.Value, out int i) && i > guidd)
				{
					guidd = i;
				}
			}
			guidd++;

			var remote_nums = remote.numbering.Root.Elements(DocxNamespace.Main + "num").ToList();
			int guidd2 = 0;
			foreach (var an in remote_nums)
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
						foreach (var numId in remote_mainDoc.Descendants(DocxNamespace.Main + "numId"))
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
								a2.SetValue(guidd);
						}

						guidd2++;
					}
				}

				guidd++;
			}

			// Checking whether there were more than 0 elements, helped me get rid of exceptions thrown while using InsertDocument
			if (numbering.Root.Elements(DocxNamespace.Main + "abstractNum").Any())
				numbering.Root.Elements(DocxNamespace.Main + "abstractNum").Last().AddAfterSelf(remote_abstractNums);

			if (numbering.Root.Elements(DocxNamespace.Main + "num").Any())
				numbering.Root.Elements(DocxNamespace.Main + "num").Last().AddAfterSelf(remote_nums);
		}

		private void MergeStyles(XDocument remote_mainDoc, DocX remote, XDocument remote_footnotes, XDocument remote_endnotes)
		{
			var local_styles = new Dictionary<string, string>();
			foreach (XElement local_style in styles.Root.Elements(DocxNamespace.Main + "style"))
			{
				XElement temp = new XElement(local_style);
				XAttribute styleId = temp.Attribute(DocxNamespace.Main + "styleId");
				string value = styleId.Value;
				styleId.Remove();
				string key = Regex.Replace(temp.ToString(), @"\s+", "");
				if (!local_styles.ContainsKey(key)) local_styles.Add(key, value);
			}

			// Add each remote style to this document.
			foreach (XElement remote_style in remote.styles.Root.Elements(DocxNamespace.Main + "style"))
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
					string local_value;
					local_styles.TryGetValue(key, out local_value);

					// If the styleIds are the same then nothing needs to be done.
					if (local_value == value)
						continue;

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
				styles.Root.Add(remote_style);
			}
		}

		private void SetMarginAttribute(XName xName, float value)
		{
			XElement body = mainDoc.Root.Element(DocxNamespace.Main + Text_Body);
			XElement sectPr = body.Element(DocxNamespace.Main + Text_SectionProperties);
			if (sectPr != null)
			{
				XElement pgMar = sectPr.Element(DocxNamespace.Main + Text_PageMargins);
				if (pgMar != null)
				{
					XAttribute top = pgMar.Attribute(xName);
					top?.SetValue(value * 20);
				}
			}
		}
	}
}