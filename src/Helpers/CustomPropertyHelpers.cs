using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    public static class CustomPropertyHelpers
    {
        internal static IReadOnlyDictionary<string, object> Get(Package packageOwner)
        {
            if (packageOwner.PartExists(Relations.CustomProperties.Uri))
            {
                var customPropDoc = packageOwner.GetPart(Relations.CustomProperties.Uri).Load();

                // Get all of the custom properties in this document
                return (
                    from p in customPropDoc.Descendants(Namespace.CustomPropertiesSchema + "property")
                    let name = p.AttributeValue("name")
                    let type = p.Descendants().Single().Name.LocalName
                    let value = p.Descendants().Single().Value
                    select new { name, val=CreateObject(type,value) }
                ).ToDictionary(p => p.name, p=> p.val, StringComparer.CurrentCultureIgnoreCase);
            }

            return new Dictionary<string, object>();
        }

        private static object CreateObject(string type, string value)
        {
            switch (type)
            {
                case CustomProperty.BOOL:
                    return value?.ToLower() == "true" || value?.ToLower() == "1";
                case CustomProperty.I4:
                {
                    return int.TryParse(value, out var result) ? result : 0;
                }
                case CustomProperty.R8:
                {
                    return double.TryParse(value, out var result) ? result : 0;
                }
                case CustomProperty.FILETIME:
                {
                    return DateTime.TryParse(value, out var result) ? result : DateTime.MinValue;
                }
                default:
                    return value ?? "";
            }
        }

        internal static bool Add(Package package, string name, string type, object value)
        {
            if (package == null)
                throw new ObjectDisposedException("Document has been disposed.");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
            if (value == null)
                throw new ArgumentException("Value cannot be null.", nameof(value));

            PackagePart customPropertiesPart;
            XDocument customPropDoc;

            // If this document does not contain a custom properties section create one.
            if (!package.PartExists(Relations.CustomProperties.Uri))
            {
                customPropertiesPart = package.CreatePart(Relations.CustomProperties.Uri, Relations.CustomProperties.ContentType, CompressionOption.Maximum);
                customPropDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                    new XElement(Namespace.CustomPropertiesSchema + "Properties", new XAttribute(XNamespace.Xmlns + "vt", Namespace.CustomVTypesSchema))
                );

                customPropertiesPart.Save(customPropDoc);
                package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, Relations.CustomProperties.RelType);
            }
            else
            {
                customPropertiesPart = package.GetPart(Relations.CustomProperties.Uri);
                customPropDoc = customPropertiesPart.Load();
            }

            if (customPropDoc.Root!.Name != Namespace.CustomPropertiesSchema + "Properties")
            {
                throw new InvalidOperationException(
                    "Custom property XML document should start with root 'Properties' element.");
            }

            // Check if a custom property already exists with this name - if so, remove it.
            var existingProperty = customPropDoc.LocalNameDescendants("property")
                    .SingleOrDefault(p => p.AttributeValue(Name.NameId)
                    .Equals(name, StringComparison.InvariantCultureIgnoreCase));

            bool exists = existingProperty != null;
            existingProperty?.Remove();

            // Get the next property id in the document
            var pid = customPropDoc.LocalNameDescendants("property")
                .Select(p => int.TryParse(p.AttributeValue(Namespace.Main + "pid"), out int result) ? result : 0)
                .DefaultIfEmpty().Max() + 1;

            // Custom props start at 2
            pid = Math.Max(2, pid);

            customPropDoc.Root!.Add(
                new XElement(Namespace.CustomPropertiesSchema + "property",
                    new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
                    new XAttribute("pid", pid),
                    new XAttribute("name", name),
                        new XElement(Namespace.CustomVTypesSchema + type, value)
                )
            );

            customPropertiesPart.Save(customPropDoc);

            return exists;
        }
    }
}
