using System.Reflection;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.XPath;

namespace DXPlus.Internal;

/// <summary>
/// Internal helpers to work with XElement/XDocument types.
/// </summary>
internal static class XLinqExtensions
{
    /// <summary>
    /// Returns whether this Xml fragment is in a document.
    /// </summary>
    public static bool InDom(this XNode? node) => node?.Parent != null;

    /// <summary>
    /// Retrieves a specific attribute value by following a path of XNames
    /// </summary>
    /// <param name="xml">Root XML element to start with</param>
    /// <param name="path">Path to follow - the final XName should be the attribute name</param>
    /// <returns>String value of the attribute, or null if any part of the path doesn't exist.</returns>
    public static string? AttributeValue(this XContainer? xml, params XName[] path)
    {
        if (xml == null) return null;
        if (path == null || path.Length == 0)
            throw new ArgumentException("Must supply the path to follow.", nameof(path));

        if (path.Length > 1)
        {
            for (int i = 0; i < path.Length - 1 && xml != null; i++)
                xml = xml.Element(path[i]);
        }

        return (xml as XElement)?.Attribute(path[^1])?.Value;
    }

    /// <summary>
    /// Look for a parent by name
    /// </summary>
    /// <param name="startAt">Node to start at</param>
    /// <param name="lookFor">Node name to look for</param>
    /// <returns>Element if found, null if none</returns>
    public static XElement? FindParent(this XElement? startAt, XName lookFor)
    {
        while (startAt != null && startAt.Name != lookFor)
            startAt = startAt.Parent;
        
        return startAt;
    }

    /// <summary>
    /// Retrieve the previous sibling to an element by name.
    /// </summary>
    /// <param name="startAt">Node to start at</param>
    /// <param name="lookFor">Node name to look for</param>
    /// <returns>Element if found, null if none</returns>
    public static XElement? PreviousSibling(this XElement? startAt, XName lookFor)
    {
        if (startAt == null) return null;

        var previousNode = startAt.PreviousNode;
        while (previousNode != null)
        {
            if (previousNode is XElement xe && xe.Name == lookFor)
                return xe;

            previousNode = previousNode.PreviousNode;
        }

        return null;
    }

    /// <summary>
    /// Retrieves a specific element by walking a path.
    /// </summary>
    /// <param name="xml">Root XML element to start with</param>
    /// <param name="path">Path to follow</param>
    /// <returns>String value of the attribute, or null if any part of the path doesn't exist.</returns>
    public static XElement? Element(this XContainer? xml, params XName[] path)
    {
        if (xml == null) return null;
        if (path == null || path.Length == 0)
            throw new ArgumentException("Must supply the path to follow.", nameof(path));

        // Walk the elements
        for (int i = 0; i < path.Length && xml != null; i++)
            xml = xml.Element(path[i]);

        return (XElement?) xml;
    }

    /// <summary>
    /// Retrieves a set of elements by walking a path.
    /// </summary>
    /// <param name="xml">Root XML element to start with</param>
    /// <param name="path">Path to follow</param>
    /// <returns>String value of the attribute, or null if any part of the path doesn't exist.</returns>
    public static IEnumerable<XElement> Elements(this XContainer? xml, params XName[] path)
    {
        if (path == null || path.Length == 0)
            throw new ArgumentException("Must supply the path to follow.", nameof(path));

        if (xml == null)
            return Enumerable.Empty<XElement>();

        // Walk the elements
        for (int i = 0; i < path.Length && xml != null; i++)
        {
            xml = xml.Element(path[i]);
        }

        return (xml as XElement)?.Elements() ?? Enumerable.Empty<XElement>();
    }

    /// <summary>
    /// Gets or creates an element based on a path.
    /// </summary>
    /// <param name="node">Starting container element</param>
    /// <param name="path">Path to create/follow</param>
    /// <returns>Final node created</returns>
    public static XElement GetOrAddElement(this XContainer node, params XName[] path)
    {
        if (node == null)
            throw new ArgumentNullException(nameof(node));

        if (path == null || path.Length == 0)
            throw new ArgumentException("Must supply the path to follow.", nameof(path));

        foreach (XName name in path)
        {
            var child = node.Element(name);
            if (child == null)
            {
                child = new XElement(name);
                node.Add(child);
            }
            node = child;
        }

        return (XElement)node;
    }

    /// <summary>
    /// Gets or creates an element based on a path.
    /// </summary>
    /// <param name="node">Starting container element</param>
    /// <param name="path">Path to create/follow</param>
    /// <returns>Final node created</returns>
    public static XElement GetOrInsertElement(this XContainer node, params XName[] path)
    {
        if (node == null)
            throw new ArgumentNullException(nameof(node));

        if (path == null || path.Length == 0)
            throw new ArgumentException("Must supply the path to follow.", nameof(path));

        foreach (XName name in path)
        {
            var child = node.Element(name);
            if (child == null)
            {
                child = new XElement(name);
                node.AddFirst(child);
            }
            node = child;
        }

        return (XElement)node;
    }

    /// <summary>
    /// Adds, removes or modifies the specified element and sets the Main:val attribute to the specified value.
    /// </summary>
    /// <param name="node">Parent node</param>
    /// <param name="name">Name of the element to add/remove</param>
    /// <param name="value">Value for the Main:val attribute</param>
    /// <returns>Created or located element</returns>
    public static XElement? AddElementVal(this XElement node, XName name, object? value)
    {
        if (value == null)
        {
            node.Element(name)?.Remove();
            return null;
        }

        var e = node.Element(name);
        if (e == null)
        {
            e = new XElement(name);
            node.Add(e);
        }
        e.SetAttributeValue(Name.MainVal, value.ToString());
        return e;
    }

    /// <summary>
    /// Gets or creates an element path + attribute based on a path.
    /// </summary>
    /// <param name="node">Starting container element</param>
    /// <param name="name">Attribute name or starting path</param>
    /// <param name="pathAndValue">Remaining path + attribute value</param>
    /// <returns>Attribute located or created</returns>
    public static XAttribute? SetAttributeValue(this XElement node, XName name, params object[] pathAndValue)
    {
        if (node == null) throw new ArgumentNullException(nameof(node));
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (pathAndValue == null || pathAndValue.Length < 1)
            throw new ArgumentException("Must include a value for the attribute.", nameof(pathAndValue));

        if (pathAndValue.Length == 1)
        {
            // This is just the attribute value -- can be null to kill the attribute.
            node.SetAttributeValue(name, pathAndValue[0]);
            return node.Attribute(name);
        }

        // Use the first name as an element.
        node = node.GetOrAddElement(name);

        XName part; object val; int index;

        // Use all but the last two elements -- that's always the attrName + value.
        for (index = 0; index < pathAndValue.Length - 2; index++)
        {
            val = pathAndValue[index];
            part = val switch
            {
                XName xn => xn,
                string sn => sn,
                _ => throw new ArgumentException($"Path cannot include {val.GetType().Name} types.", nameof(pathAndValue)),
            };
            node = node.GetOrAddElement(part);
        }

        val = pathAndValue[index++];
        part = val switch
        {
            XName xn => xn,
            string sn => sn,
            _ => throw new ArgumentException($"Path cannot include {val.GetType().Name} types.", nameof(pathAndValue)),
        };

        node.SetAttributeValue(part, pathAndValue[index]);
        return node.Attribute(part);
    }

    /// <summary>
    /// Find an element by an attribute value
    /// </summary>
    /// <param name="nodes">Nodes to search</param>
    /// <param name="name">Attribute name to look for</param>
    /// <param name="attributeValue">Value to match</param>
    /// <returns></returns>
    public static XElement? FindByAttrVal(this IEnumerable<XElement> nodes, XName name, string attributeValue) 
        => nodes.FirstOrDefault(node => node.AttributeValue(name)?.Equals(attributeValue) == true);

    /// <summary>
    /// Retrieve the "val" or "w:val" attribute from an element tag.
    /// </summary>
    /// <param name="el">Element to examine</param>
    /// <returns>Attribute object, or null if the XML attribute is missing.</returns>
    public static XAttribute? GetValAttr(this XElement? el)
    {
        if (el == null) return null;
        var valAttr = el.Attribute(Name.MainVal.LocalName);
        return valAttr ?? el.Attribute(Name.MainVal);
    }

    /// <summary>
    /// Retrieves the text value of the "val" or "w:val" attribute on an element tag.
    /// </summary>
    /// <param name="el">Element to examine</param>
    /// <param name="defaultValue">Value to return if the attribute does not exist.</param>
    /// <returns>Value of the attribute, or the default value if it doesn't exist.</returns>
    public static string? GetVal(this XElement? el, string? defaultValue = "") 
        => GetValAttr(el)?.Value ?? defaultValue;

    /// <summary>
    /// Retrieves the text value of the specified attribute name on an element tag.
    /// </summary>
    /// <param name="el">Element to examine</param>
    /// <param name="name">Name to look for</param>
    /// <param name="defaultValue">Value to return if attribute is missing</param>
    /// <returns>Value of the attribute, or the default</returns>
    public static string? AttributeValue(this XElement? el, XName name, string? defaultValue = "")
    {
        var attr = el?.Attribute(name);
        return attr != null ? attr.Value : defaultValue;
    }

    /// <summary>
    /// Retrieves a true/false boolean value for an attribute on the specified element tag.
    /// </summary>
    /// <param name="el">Element to examine</param>
    /// <param name="name">Name of the attribute</param>
    /// <param name="defaultValue">Default value, defaults to false</param>
    /// <returns>Value of the boolean attribute, or the default value</returns>
    public static bool? BoolAttributeValue(this XElement? el, XName name, bool? defaultValue = false)
    {
        if (name == null) throw new ArgumentNullException(nameof(name));
        var attr = el?.Attribute(name);
        if (attr == null) return defaultValue;

        var val = attr.Value.Trim();
        return string.Equals(val, "true", StringComparison.OrdinalIgnoreCase) 
               || string.Equals(val, "on", StringComparison.OrdinalIgnoreCase)
               || val == "1";
    }

    /// <summary>
    /// Returns all elements in the given XML container that match a specific name ignoring namespaces.
    /// </summary>
    /// <param name="xml">XML container</param>
    /// <param name="localName">Name to look for</param>
    /// <returns>All matching elements</returns>
    public static IEnumerable<XElement> LocalNameElements(this XContainer xml, string localName)
    {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        if (localName == null) throw new ArgumentNullException(nameof(localName));
        return xml.Elements().Where(e => e.Name.LocalName.Equals(localName));
    }

    /// <summary>
    /// Returns the first matching element in the given container by name ignoring namespaces.
    /// </summary>
    /// <param name="e">XML container</param>
    /// <param name="localName">Name to look for</param>
    /// <returns>Matching element, or null</returns>
    public static XElement? FirstLocalNameDescendant(this XContainer e, string localName)
    {
        if (e == null) throw new ArgumentNullException(nameof(e));
        if (localName == null) throw new ArgumentNullException(nameof(localName));
        return e.LocalNameDescendants(localName).FirstOrDefault();
    }

    /// <summary>
    /// Returns all matching descendent elements in the given container by name ignoring namespaces.
    /// </summary>
    /// <param name="xml">XML container</param>
    /// <param name="localName">Name to look for</param>
    /// <returns>Matching elements</returns>
    public static IEnumerable<XElement> LocalNameDescendants(this XContainer? xml, string localName)
    {
        if (string.IsNullOrWhiteSpace(localName))
            throw new ArgumentNullException(nameof(localName));

        if (xml == null)
            yield break;

        if (localName.Contains('/'))
        {
            int pos = localName.IndexOf('/');
            string name = localName[..pos];
            localName = localName[(pos + 1)..];

            foreach (var item in xml.Descendants().Where(e => e.Name.LocalName == name))
            {
                foreach (var child in LocalNameDescendants(item, localName))
                    yield return child;
            }
        }
        else
        {
            foreach (var item in xml.Descendants().Where(e => e.Name.LocalName == localName))
                yield return item;
        }
    }

    /// <summary>
    /// Get value from XElement and convert it to enum
    /// </summary>
    /// <typeparam name="T">Enum type</typeparam>
    public static T GetEnumValue<T>(this XElement element) where T : Enum
    {
        if (element == null) throw new ArgumentNullException(nameof(element));
        if (!TryGetEnumValue(element, out T? result))
            throw new ArgumentException($"{element.GetVal()} could not be matched to enum {typeof(T).Name}.");
        
        return result!;
    }

    /// <summary>
    /// Get value from XElement and convert it to enum
    /// </summary>
    /// <typeparam name="T">Enum type</typeparam>
    public static T GetEnumValue<T>(this XAttribute attr) where T : Enum
    {
        if (attr == null) throw new ArgumentNullException(nameof(attr));
        if (!TryGetEnumValue(attr, out T? result))
            throw new ArgumentException($"{attr.Value} could not be matched to enum {typeof(T).Name}.");
        
        return result!;
    }

    /// <summary>
    /// Convert value to xml string and set it into XElement
    /// </summary>
    /// <typeparam name="T">Enum type</typeparam>
    public static void SetEnumValue<T>(this XElement element, T value) where T : Enum
    {
        if (element == null) throw new ArgumentNullException(nameof(element));
        element.SetAttributeValue("val", GetEnumName(value));
    }

    /// <summary>
    /// Convert an attribute to an enumeration
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="attr"></param>
    /// <param name="result"></param>
    /// <returns></returns>
    public static bool TryGetEnumValue<T>(this XAttribute? attr, out T? result) where T : Enum
    {
        if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
        {
            string value = attr.Value;
            foreach (T e in Enum.GetValues(typeof(T)))
            {
                var valueName = e.ToString();
                var fi = typeof(T).GetField(valueName);
                if (fi != null)
                {
                    string name = fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? valueName;
                    if (string.Equals(name, value, StringComparison.OrdinalIgnoreCase))
                    {
                        result = e;
                        return true;
                    }
                }
            }
        }

        result = default;
        return false;
    }

    /// <summary>
    /// Convert a text string to an enum value
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="value"></param>
    /// <param name="result"></param>
    /// <returns></returns>
    public static bool TryGetEnumValue<T>(this string? value, out T? result) where T : Enum
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            foreach (T e in Enum.GetValues(typeof(T)))
            {
                var fi = typeof(T).GetField(e.ToString());
                string name = fi?.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? e.ToString();
                if (string.Equals(name, value, StringComparison.OrdinalIgnoreCase))
                {
                    result = e;
                    return true;
                }
            }
        }

        result = default;
        return false;
    }

    /// <summary>
    /// Convert the val() attribute of an Element to an enumeration
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="element"></param>
    /// <param name="result"></param>
    /// <returns></returns>
    public static bool TryGetEnumValue<T>(this XElement element, out T? result) where T : Enum
    {
        if (element == null) throw new ArgumentNullException(nameof(element));
        return TryGetEnumValue(element.GetVal(), out result);
    }

    /// <summary>
    /// Return XML string for this value
    /// </summary>
    /// <typeparam name="T">Enum type</typeparam>
    public static string GetEnumName<T>(this T value) where T : Enum
    {
        var fi = typeof(T).GetField(value.ToString());
        return fi?.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? value.ToCamelCase();
    }

    /// <summary>
    /// Normalize an XML element graph by recursively ordering attribute values and child elements
    /// based on name and value.
    /// </summary>
    /// <param name="element">Element to normalize</param>
    /// <returns>New copy of element with ordered attributes and children</returns>
    public static XElement Normalize(this XElement element)
    {
        if (element.HasElements)
        {
            return new XElement(
                element.Name,
                element.Attributes().OrderBy(a => a.Name.ToString()),
                element.Elements()
                    .Select(Normalize)
                    .OrderBy(e => e.Name.ToString())
                    .ThenBy(e => e.Attributes().Count())
                    .ThenBy(e => string.Join(',', e.Attributes().OrderBy(a => a.Name.ToString()).Select(a => $"{a.Name}:{a.Value}")))
                    .ThenBy(e => e.Value));
        }

        if (element.IsEmpty || string.IsNullOrEmpty(element.Value))
        {
            return new XElement(element.Name,
                element.Attributes().OrderBy(a => a.Name.ToString()));
        }

        return new XElement(element.Name,
            element.Attributes().OrderBy(a => a.Name.ToString()),
            element.Value);
    }

    /// <summary>
    /// Performs an XPath query against an element with the proper Word document namespaces.
    /// </summary>
    /// <param name="element">Element</param>
    /// <param name="query">Query</param>
    /// <returns>Results from query</returns>
    public static IEnumerable<XElement> QueryElements(this XElement element, string query) 
        => element.XPathSelectElements(query, Namespace.NamespaceManager());

    /// <summary>
    /// Performs an XPath query against an element with the proper Word document namespaces.
    /// </summary>
    /// <param name="element">Element</param>
    /// <param name="query">Query</param>
    /// <returns>Results from query</returns>
    public static XElement? QueryElement(this XElement element, string query) 
        => element.XPathSelectElement(query, Namespace.NamespaceManager());
}

