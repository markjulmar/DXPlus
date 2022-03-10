using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus.Helpers;

/// <summary>
/// Helper methods to work with custom document properties.
/// </summary>
public static class CustomPropertyHelpers
{
    /// <summary>
    /// Returns all the properties defined in a specific document package.
    /// </summary>
    /// <param name="packageOwner">Package owner</param>
    /// <returns>Dictionary of properties</returns>
    internal static IReadOnlyDictionary<string, object> Get(Package packageOwner)
    {
        if (packageOwner.PartExists(Relations.CustomProperties.Uri))
        {
            var customPropDoc = packageOwner.GetPart(Relations.CustomProperties.Uri).Load();
            return customPropDoc.Descendants(Namespace.CustomPropertiesSchema + "property")
                .Select(element => new {
                    name = element.AttributeValue("name") ?? throw new DocumentFormatException("CustomProperties"),
                    value = CreateObject(element.Descendants().Single().Name.LocalName, element.Descendants().Single().Value)})
                .ToDictionary(p => p.name, p => p.value, StringComparer.CurrentCultureIgnoreCase);
        }
        return new Dictionary<string, object>();
    }

    /// <summary>
    /// Returns a discrete type based on the custom property value.
    /// </summary>
    /// <param name="type">Object type</param>
    /// <param name="value">Object value</param>
    /// <returns>Discrete object</returns>
    private static object CreateObject(string type, string value)
    {
        return type switch
        {
            CustomProperty.BOOL => value.ToLower() == "true" || value.ToLower() == "1",
            CustomProperty.I4 => int.TryParse(value, out var iv) ? iv : 0,
            CustomProperty.R8 => double.TryParse(value, out var rv) ? rv : 0,
            CustomProperty.FILETIME => DateTime.TryParse(value, out var dt) ? dt : DateTime.MinValue,
            _ => value
        };
    }

    /// <summary>
    /// Adds a new custom property to the document.
    /// </summary>
    /// <param name="package">Package to add custom value to</param>
    /// <param name="name">Name of property</param>
    /// <param name="type">Type of property</param>
    /// <param name="value">Value of property</param>
    /// <returns>True/False if property was successfully added</returns>
    /// <exception cref="ObjectDisposedException"></exception>
    /// <exception cref="ArgumentException"></exception>
    /// <exception cref="InvalidOperationException"></exception>
    internal static bool Add(Package package, string name, string type, object value)
    {
        if (package == null) throw new ObjectDisposedException("Document has been disposed.");
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));
        if (value == null) throw new ArgumentException("Value cannot be null.", nameof(value));

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
            .SingleOrDefault(p => p.AttributeValue(Name.NameId)?
            .Equals(name, StringComparison.InvariantCultureIgnoreCase) == true);

        bool exists = existingProperty != null;
        existingProperty?.Remove();

        // Get the next property id in the document
        var pid = customPropDoc.LocalNameDescendants("property")
            .Select(p => int.TryParse(p.AttributeValue("pid"), out int result) ? result : 0)
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