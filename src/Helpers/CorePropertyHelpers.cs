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
    internal static class CorePropertyHelpers
    {
        public static IReadOnlyDictionary<DocumentPropertyName, string> Get(Package package)
        {
            if (package == null)
                throw new ObjectDisposedException("Document has been disposed.");

            var values = new Dictionary<DocumentPropertyName, string>();

            if (package.PartExists(Relations.CoreProperties.Uri))
            {
                var doc = package.GetPart(Relations.CoreProperties.Uri).Load();
                foreach (var e in doc.Root!.Elements())
                {
                    var name = $"{doc.Root.GetPrefixOfNamespace(e.Name.Namespace)}:{e.Name.LocalName}";
                    if (name.TryGetEnumValue(out DocumentPropertyName result))
                    {
                        values.Add(result, e.Value);
                    }
                }
            }

            return values;
        }

        public static string SetValue(Package package, string name, string value)
        {
            if (package == null)
                throw new ObjectDisposedException("Document has been disposed.");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(value));

            XDocument corePropDoc;
            PackagePart corePropPart;

            // Create the core document if it doesn't exist yet.
            if (!package.PartExists(Relations.CoreProperties.Uri))
            {
                corePropPart = CreateCoreProperties(package, out corePropDoc);
            }
            else
            {
                corePropPart = package.GetPart(Relations.CoreProperties.Uri);
                corePropDoc = corePropPart.Load();
            }

            if (!HelperFunctions.SplitXmlName(name, out var ns, out var localName))
                ns = "cp"; // default

            var xns = corePropDoc.Root!.GetNamespaceOfPrefix(ns);
            if (xns == null)
                throw new InvalidOperationException($"Unable to identify namespace {ns} used core property {localName}.");

            var corePropElement = corePropDoc.Root!.Elements().SingleOrDefault(e => e.Name == xns + localName);
            if (corePropElement != null)
            {
                corePropElement.SetValue(value);
            }
            else
            {
                corePropDoc.Root.Add(new XElement(xns + localName, value));
            }

            corePropPart.Save(corePropDoc);
            return localName;
        }

        public static PackagePart CreateCoreProperties(Package package, out XDocument corePropDoc)
        {
            string userName = Environment.UserInteractive ? Environment.UserName : string.Empty;
            if (string.IsNullOrWhiteSpace(userName)) userName = "Office User";

            var corePropPart = package.CreatePart(Relations.CoreProperties.Uri, Relations.CoreProperties.ContentType, CompressionOption.Maximum);
            corePropDoc = Resource.CorePropsXml(userName, DateTime.UtcNow);
            Debug.Assert(corePropDoc.Root != null);

            corePropPart.Save(corePropDoc);
            package.CreateRelationship(corePropPart.Uri, TargetMode.Internal, Relations.CoreProperties.RelType);
            return corePropPart;
        }
    }
}