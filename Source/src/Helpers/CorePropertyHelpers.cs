using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Resources;

namespace DXPlus.Helpers
{
    /// <summary>
    /// Helper class to deal with DOCX core properties
    /// </summary>
    public static class CorePropertyHelpers
    {
        internal static Dictionary<string, string> Get(Package package)
        {
            if (!package.PartExists(DocxSections.DocPropsCoreUri))
                return new Dictionary<string, string>();

            // Get all of the core properties in this document
            var corePropDoc = package.GetPart(DocxSections.DocPropsCoreUri).Load();
            return corePropDoc.Root!.Elements()
                .Select(docProperty =>
                    new KeyValuePair<string, string>(
                        $"{corePropDoc.Root.GetPrefixOfNamespace(docProperty.Name.Namespace)}:{docProperty.Name.LocalName}",
                        docProperty.Value))
                .ToDictionary(p => p.Key, v => v.Value);
        }

        internal static string Add(Package package, string name, string value)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(value));

            XDocument corePropDoc;
            PackagePart corePropPart;

            // Create the core document if it doesn't exist yet.
            if (!package.PartExists(DocxSections.DocPropsCoreUri))
            {
                corePropPart = package.CreatePart(Relations.CoreProperties.Uri, Relations.CoreProperties.ContentType, CompressionOption.Maximum);
                corePropDoc = Resource.CorePropsXml(Environment.UserName, DateTime.UtcNow);
                Debug.Assert(corePropDoc.Root != null);

                corePropPart.Save(corePropDoc);
                package.CreateRelationship(corePropPart.Uri, TargetMode.Internal, Relations.CoreProperties.RelType);
            }
            else
            {
                corePropPart = package.GetPart(DocxSections.DocPropsCoreUri);
                corePropDoc = corePropPart.Load();
            }

            if (!HelperFunctions.SplitXmlName(name, out var ns, out var localName))
                ns = "cp";

            var corePropElement = corePropDoc.Root!.Elements().SingleOrDefault(e => e.Name.LocalName.Equals(localName));
            if (corePropElement != null)
            {
                corePropElement.SetValue(value);
            }
            else
            {
                var xns = corePropDoc.Root.GetNamespaceOfPrefix(ns);
                if (xns == null)
                    throw new InvalidOperationException($"Unable to identify namespace {ns} used core property {localName}.");

                corePropDoc.Root.Add(new XElement(xns + localName, value));
            }

            corePropPart.Save(corePropDoc);
            return localName;
        }
    }
}