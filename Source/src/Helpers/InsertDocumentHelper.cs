using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    partial class DocX
    {
		/// <summary>
		/// Insert the contents of another document at the end of this document.
		/// If the document being inserted contains Images, CustomProperties and or custom styles, these will be correctly inserted into the new document.
		/// In the case of Images, new ID's are generated for the Images being inserted to avoid ID conflicts. CustomProperties with the same name will be ignored not replaced.
		/// </summary>
		/// <param name="otherDoc">The document to insert at the end of this document.</param>
		/// <param name="append">If true, document is inserted at the end, otherwise document is inserted at the beginning.</param>
		private void InsertDocument(DocX otherDoc, bool append = true)
		{
			if (otherDoc == null)
				throw new ArgumentNullException(nameof(otherDoc));

			ThrowIfObjectDisposed();
			otherDoc.ThrowIfObjectDisposed();

			// Copy all the XML bits
			var otherMainDoc = new XDocument(otherDoc.mainDoc);
			var otherFootnotes = otherDoc.footnotesDoc != null ? new XDocument(otherDoc.footnotesDoc) : null;
			var otherEndnotes = otherDoc.endnotesDoc != null ? new XDocument(otherDoc.endnotesDoc) : null;
            var otherBody = otherDoc.Xml;

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
			foreach (var otherPackagePart in otherDoc.Package.GetParts())
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
			foreach (var rel in otherDoc.PackagePart.GetRelationshipsByType($"{DocxNamespace.RelatedDoc.NamespaceName}/hyperlink"))
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
			foreach (var rel in otherDoc.PackagePart.GetRelationshipsByType($"{DocxNamespace.RelatedDoc.NamespaceName}/oleObject"))
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
													$"{DocxNamespace.RelatedDoc.NamespaceName}/image");

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

        private static string ComputeHashString(Stream stream)
        {
            byte[] hash = MD5.Create().ComputeHash(stream);

            var sb = new StringBuilder();
            foreach (var value in hash)
                sb.Append(value.ToString("X2"));

            return sb.ToString();
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
			foreach (var remoteRelationship in otherDocument.PackagePart.GetRelationships())
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
	}
}
