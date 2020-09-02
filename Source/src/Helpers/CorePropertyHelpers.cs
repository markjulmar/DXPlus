using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    /// <summary>
    /// Helper class to deal with DOCX core properties
    /// </summary>
    public static class CorePropertyHelpers
    {
        internal static Dictionary<string, string> Get(Package packageOwner)
        {
            if (!packageOwner.PartExists(DocxSections.DocPropsCoreUri))
                return new Dictionary<string, string>();

            // Get all of the core properties in this document
            var corePropDoc = packageOwner.GetPart(DocxSections.DocPropsCoreUri).Load();
            return corePropDoc.Root!.Elements()
                .Select(docProperty =>
                    new KeyValuePair<string, string>(
                        $"{corePropDoc.Root.GetPrefixOfNamespace(docProperty.Name.Namespace)}:{docProperty.Name.LocalName}",
                        docProperty.Value))
                .ToDictionary(p => p.Key, v => v.Value);
        }

        internal static void Add(DocX document, string name, string value)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(value));
            if (!document.Package.PartExists(DocxSections.DocPropsCoreUri))
                throw new Exception("Core properties part doesn't exist.");

            string propertyNamespacePrefix = name.Contains(":") ? name.Split(new[] {':'})[0] : "cp";
            string propertyLocalName = name.Contains(":") ? name.Split(new[] {':'})[1] : name;

            var corePropPart = document.Package.GetPart(DocxSections.DocPropsCoreUri);
            var corePropDoc = corePropPart.Load();

            var corePropElement = corePropDoc.Root!.Elements()
                .SingleOrDefault(e => e.Name.LocalName.Equals(propertyLocalName));
            if (corePropElement != null)
            {
                corePropElement.SetValue(value);
            }
            else
            {
                var propertyNamespace = corePropDoc.Root.GetNamespaceOfPrefix(propertyNamespacePrefix);
                if (propertyNamespace == null)
                    throw new InvalidOperationException("Unable to identify namespace for core property.");
                corePropDoc.Root.Add(new XElement(Namespace.Main + propertyLocalName,
                    propertyNamespace.NamespaceName, value));
            }

            corePropPart.Save(corePropDoc);
            document?.UpdateCorePropertyUsages(propertyLocalName, value);
        }
    }
}