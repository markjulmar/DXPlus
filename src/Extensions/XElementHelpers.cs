using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DXPlus
{
    /// <summary>
    /// Internal helpers to work with XElement/XDocument types.
    /// </summary>
    internal static class XLinqExtensions
    {
        /// <summary>
        /// Retrieves a specific attribute value by following a path of XNames
        /// </summary>
        /// <param name="xml">Root XML element to start with</param>
        /// <param name="path">Path to follow - the final XName should be the attribute name</param>
        /// <returns>String value of the attribute, or null if any part of the path doesn't exist.</returns>
        public static string AttributeValue(this XContainer xml, params XName[] path)
        {
            if (xml == null)
            {
                return null;
            }

            if (path == null || path.Length == 0)
            {
                throw new ArgumentException("Must supply the path to follow.", nameof(path));
            }

            if (path.Length > 1)
            {
                for (int i = 0; i < path.Length - 1 && xml != null; i++)
                {
                    xml = xml.Element(path[i]);
                }
            }

            return (xml as XElement)?.Attribute(path[^1])?.Value;
        }

        /// <summary>
        /// Retrieves a specific element by walking a path.
        /// </summary>
        /// <param name="xml">Root XML element to start with</param>
        /// <param name="path">Path to follow</param>
        /// <returns>String value of the attribute, or null if any part of the path doesn't exist.</returns>
        public static XElement Element(this XContainer xml, params XName[] path)
        {
            if (xml == null)
            {
                return null;
            }

            if (path == null || path.Length == 0)
            {
                throw new ArgumentException("Must supply the path to follow.", nameof(path));
            }

            // Walk the elements
            for (int i = 0; i < path.Length && xml != null; i++)
            {
                xml = xml.Element(path[i]);
            }

            return (XElement)xml;
        }

        /// <summary>
        /// Retrieves a set of elements by walking a path.
        /// </summary>
        /// <param name="xml">Root XML element to start with</param>
        /// <param name="path">Path to follow</param>
        /// <returns>String value of the attribute, or null if any part of the path doesn't exist.</returns>
        public static IEnumerable<XElement> Elements(this XContainer xml, params XName[] path)
        {
            if (xml == null)
            {
                return null;
            }

            if (path == null || path.Length == 0)
            {
                throw new ArgumentException("Must supply the path to follow.", nameof(path));
            }

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
            {
                throw new ArgumentNullException(nameof(node));
            }

            if (path == null || path.Length == 0)
            {
                throw new ArgumentException("Must supply the path to follow.", nameof(path));
            }

            foreach (XName name in path)
            {
                XElement child = node.Element(name);
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
            {
                throw new ArgumentNullException(nameof(node));
            }

            if (path == null || path.Length == 0)
            {
                throw new ArgumentException("Must supply the path to follow.", nameof(path));
            }

            foreach (XName name in path)
            {
                XElement child = node.Element(name);
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
        public static XElement AddElementVal(this XElement node, XName name, object value)
        {
            if (value == null)
            {
                node.Element(name)?.Remove();
                return null;
            }

            XElement e = node.Element(name);
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
        public static XAttribute SetAttributeValue(this XElement node, XName name, params object[] pathAndValue)
        {
            if (node == null)
            {
                throw new ArgumentNullException(nameof(node));
            }

            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            if (pathAndValue == null || pathAndValue.Length < 1)
            {
                throw new ArgumentException("Must include a value for the attribute.", nameof(pathAndValue));
            }

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
            for (index = 0; index < pathAndValue.Length - 2 && node != null; index++)
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

            node!.SetAttributeValue(part, pathAndValue[index]);
            return node.Attribute(part);
        }

        /// <summary>
        /// Find an element by an attribute value
        /// </summary>
        /// <param name="nodes">Nodes to search</param>
        /// <param name="name">Attribute name to look for</param>
        /// <param name="attributeValue">Value to match</param>
        /// <returns></returns>
        public static XElement FindByAttrVal(this IEnumerable<XElement> nodes, XName name, string attributeValue)
        {
            return nodes?.FirstOrDefault(node => node.AttributeValue(name).Equals(attributeValue));
        }

        public static XAttribute GetValAttr(this XElement el)
        {
            if (el == null)
            {
                return null;
            }

            XAttribute valAttr = el.Attribute("val");
            return valAttr ?? el.Attribute(Name.MainVal);
        }

        public static string GetVal(this XElement el, string defaultValue = "")
        {
            return GetValAttr(el)?.Value ?? defaultValue;
        }

        public static string AttributeValue(this XElement el, XName name, string defaultValue = "")
        {
            XAttribute attr = el?.Attribute(name);
            return attr != null ? attr.Value : defaultValue;
        }

        public static bool BoolAttributeValue(this XElement el, XName name, bool defaultValue = false)
        {
            XAttribute attr = el?.Attribute(name);
            if (attr == null)
            {
                return defaultValue;
            }

            string val = attr.Value.Trim();
            return string.Equals(val, "true", StringComparison.OrdinalIgnoreCase) || val == "1";
        }

        public static IEnumerable<XElement> LocalNameElements(this XContainer xml, string localName)
        {
            return xml.Elements().Where(e => e.Name.LocalName.Equals(localName));
        }

        public static XElement FirstLocalNameDescendant(this XContainer e, string localName)
        {
            return e.LocalNameDescendants(localName).FirstOrDefault();
        }

        public static IEnumerable<XElement> LocalNameDescendants(this XContainer xml, string localName)
        {
            if (xml == null)
            {
                yield break;
            }

            if (string.IsNullOrWhiteSpace(localName))
            {
                throw new ArgumentNullException(nameof(localName));
            }

            if (localName.Contains('/'))
            {
                int pos = localName.IndexOf('/');
                string name = localName.Substring(0, pos);
                localName = localName.Substring(pos + 1);

                foreach (XElement item in xml.Descendants().Where(e => e.Name.LocalName == name))
                {
                    foreach (XElement child in LocalNameDescendants(item, localName))
                    {
                        yield return child;
                    }
                }
            }
            else
            {
                foreach (XElement item in xml.Descendants().Where(e => e.Name.LocalName == localName))
                {
                    yield return item;
                }
            }
        }

        public static IEnumerable<XAttribute> DescendantAttributes(this XContainer xml, XName attribName)
        {
            return xml.Descendants().Attributes(attribName);
        }

        public static IEnumerable<string> DescendantAttributeValues(this XElement xml, XName attribName)
        {
            return xml.DescendantAttributes(attribName).Select(a => a.Value);
        }

        /// <summary>
        /// Get value from XElement and convert it to enum
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        public static T GetEnumValue<T>(this XElement element)
        {
            if (element == null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (!TryGetEnumValue(element, out T result))
            {
                throw new ArgumentException($"{element.GetVal()} could not be matched to enum {typeof(T).Name}.");
            }

            return result;
        }

        /// <summary>
        /// Get value from XElement and convert it to enum
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        public static T GetEnumValue<T>(this XAttribute attr)
        {
            if (!TryGetEnumValue(attr, out T result))
            {
                throw new ArgumentException($"{attr.Value} could not be matched to enum {typeof(T).Name}.");
            }

            return result;
        }

        /// <summary>
        /// Convert value to xml string and set it into XElement
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        public static void SetEnumValue<T>(this XElement element, T value)
        {
            if (element == null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            element.SetAttributeValue("val", GetEnumName(value));
        }

        /// <summary>
        /// Convert an attribute to an enumeration
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="attr"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public static bool TryGetEnumValue<T>(this XAttribute attr, out T result)
        {
            if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
            {
                string value = attr.Value;
                foreach (T e in Enum.GetValues(typeof(T)))
                {
                    FieldInfo fi = typeof(T).GetField(e.ToString());
                    string name = fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? e.ToString();
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
        /// Convert a text string to an enum value
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public static bool TryGetEnumValue<T>(this string value, out T result)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                foreach (T e in Enum.GetValues(typeof(T)))
                {
                    FieldInfo fi = typeof(T).GetField(e.ToString());
                    string name = fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? e.ToString();
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
        public static bool TryGetEnumValue<T>(this XElement element, out T result)
        {
            return TryGetEnumValue(element?.GetVal(), out result);
        }

        /// <summary>
        /// Return xml string for this value
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        public static string GetEnumName<T>(this T value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            FieldInfo fi = typeof(T).GetField(value.ToString());
            return fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? value.ToCamelCase();
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
    }
}