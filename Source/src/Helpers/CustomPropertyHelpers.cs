using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
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
                    from p in customPropDoc.Descendants(Namespace.CustomPropertiesSchema + "property")
                    let name = p.AttributeValue(Name.NameId)
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
					new XElement(Namespace.CustomPropertiesSchema + "Properties",
						new XAttribute(XNamespace.Xmlns + "vt", Namespace.CustomVTypesSchema)
					)
				);

				customPropertiesPart.Save(customPropDoc);
                document.Package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, $"{Namespace.RelatedDoc.NamespaceName}/custom-properties");
			}
			else
			{
				customPropertiesPart = document.Package.GetPart(DocxSections.DocPropsCustom);
				customPropDoc = customPropertiesPart.Load();
			}

			// Get the next property id in the document
			var pid = customPropDoc.LocalNameDescendants("property")
				.Select(p => int.TryParse(p.AttributeValue(Namespace.Main + "pid"), out int result) ? result : 0)
				.DefaultIfEmpty().Max() + 1;

			// Check if a custom property already exists with this name - if so, remove it.
			customPropDoc.LocalNameDescendants("property")
					.SingleOrDefault(p => p.AttributeValue(Name.NameId)
					.Equals(property.Name, StringComparison.InvariantCultureIgnoreCase))
					?.Remove();

			var propertiesElement = customPropDoc.Element(Namespace.CustomPropertiesSchema + "Properties");
			propertiesElement.Add(
				new XElement(Namespace.CustomPropertiesSchema + "property",
					new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
					new XAttribute("pid", pid),
					new XAttribute("name", property.Name),
						new XElement(Namespace.CustomVTypesSchema + property.Type, property.Value ?? string.Empty)
				)
			);

			// Save the custom properties
			customPropertiesPart.Save(customPropDoc);

			// Refresh the places using this property.
            document?.UpdateCustomPropertyUsages(property);
        }
	}
}
