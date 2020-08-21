using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    public static class CustomPropertyHelpers
    {
        internal static Dictionary<string, CustomProperty> Get(Package packageOwner)
        {
            if (packageOwner.PartExists(DocxSections.DocPropsCustom))
            {
                var customPropDoc = packageOwner.GetPart(DocxSections.DocPropsCustom).Load();

                // Get all of the custom properties in this document
                return (
                    from p in customPropDoc.Descendants(DocxNamespace.CustomPropertiesSchema + "property")
                    let name = p.AttributeValue(DocxNamespace.Main + "name")
                    let type = p.Descendants().Single().Name.LocalName
                    let value = p.Descendants().Single().Value
                    select new CustomProperty(name, type, value)
                ).ToDictionary(p => p.Name, StringComparer.CurrentCultureIgnoreCase);
            }

            return new Dictionary<string, CustomProperty>();
        }

        internal static void Add(DocX document, CustomProperty property)
        {
			PackagePart customPropertiesPart;
			XDocument customPropDoc;

			// If this document does not contain a custom properties section create one.
			if (!document.Package.PartExists(DocxSections.DocPropsCustom))
			{
				customPropertiesPart = document.Package.CreatePart(new Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml", CompressionOption.Maximum);
				customPropDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
					new XElement(DocxNamespace.CustomPropertiesSchema + "Properties",
						new XAttribute(XNamespace.Xmlns + "vt", DocxNamespace.CustomVTypesSchema)
					)
				);

				customPropertiesPart.Save(customPropDoc);
                document.Package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, $"{DocxNamespace.RelatedDoc.NamespaceName}/custom-properties");
			}
			else
			{
				customPropertiesPart = document.Package.GetPart(DocxSections.DocPropsCustom);
				customPropDoc = customPropertiesPart.Load();
			}

			// Get the next property id in the document
			var pid = customPropDoc.LocalNameDescendants("property")
				.Select(p => int.TryParse(p.AttributeValue(DocxNamespace.Main + "pid"), out int result) ? result : 0)
				.DefaultIfEmpty().Max() + 1;

			// Check if a custom property already exists with this name - if so, remove it.
			customPropDoc.LocalNameDescendants("property")
					.SingleOrDefault(p => p.AttributeValue(DocxNamespace.Main + "name")
					.Equals(property.Name, StringComparison.InvariantCultureIgnoreCase))
					?.Remove();

			var propertiesElement = customPropDoc.Element(DocxNamespace.CustomPropertiesSchema + "Properties");
			propertiesElement.Add(
				new XElement(DocxNamespace.CustomPropertiesSchema + "property",
					new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
					new XAttribute("pid", pid),
					new XAttribute("name", property.Name),
						new XElement(DocxNamespace.CustomVTypesSchema + property.Type, property.Value ?? string.Empty)
				)
			);

			// Save the custom properties
			customPropertiesPart.Save(customPropDoc);

			// Refresh the places using this property.
            Update(document, property);
        }

        /// <summary>
        /// Update the custom properties inside the document
        /// </summary>
        /// <param name="document">The DocX document</param>
        /// <param name="property">Custom property</param>
        /// <remarks>Different version of Word create different Document XML.</remarks>
        internal static void Update(DocX document, CustomProperty property)
		{
			if (document == null)
				throw new ArgumentNullException(nameof(document));
			if (property == null)
				throw new ArgumentNullException(nameof(property));

			var documents = new List<XElement> { document.mainDoc.Root };
            var value = property.Value?.ToString() ?? string.Empty;

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

			string matchCustomPropertyName = property.Name;
			if (property.Name.Contains(" "))
				matchCustomPropertyName = "\"" + property.Name + "\"";

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
								var ele = (XElement)nextNode;
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


	}
}
