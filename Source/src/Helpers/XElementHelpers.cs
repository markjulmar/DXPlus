using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DXPlus.Helpers
{
    internal static class XElementHelpers
    {
        /// <summary>
        /// Get the rPr element from a parent, or create it if it doesn't exist.
        /// </summary>
        /// <param name="owner"></param>
        /// <param name="create">True to create it if it doesn't exist</param>
        /// <returns></returns>
        internal static XElement GetRunProps(this XElement owner, bool create = true)
        {
            var rPr = owner.Element(DocxNamespace.Main + "rPr");
            if (rPr == null && create)
            {
                rPr = new XElement(DocxNamespace.Main + "rPr");
                owner.AddFirst(rPr);
            }

            return rPr;
        }

        /// <summary>
        /// If a text element or delText element, starts or ends with a space,
        /// it must have the attribute space, otherwise it must not have it.
        /// </summary>
        /// <param name="e">The (t or delText) element check</param>
        public static XElement PreserveSpace(this XElement e)
        {
            if (!e.Name.Equals(DocxNamespace.Main + "t")
                && !e.Name.Equals(DocxNamespace.Main + "delText"))
            {
                throw new ArgumentException($"{nameof(PreserveSpace)} can only work with elements of type 't' or 'delText'", nameof(e));
            }

            // Check if this w:t contains a space attribute
            var space = e.Attributes().SingleOrDefault(a => a.Name.Equals(XNamespace.Xml + "space"));

            // This w:t's text begins or ends with whitespace
            if (e.Value.StartsWith(" ") || e.Value.EndsWith(" "))
            {
                // If this w:t contains no space attribute, add one.
                if (space == null)
                {
                    e.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                }
            }

            // This w:t's text does not begin or end with a space
            else
            {
                // If this w:r contains a space attribute, remove it.
                space?.Remove();
            }

            return e;
        }

        public static XElement GetOrCreateElement(this XContainer el, XName name, string defaultValue = "")
        {
            if (el == null)
                throw new ArgumentNullException(nameof(el));
            
            var node = el.Element(name);
            if (node == null)
            {
                node = new XElement(name, defaultValue);
                el.Add(node);
            }
            return node;
        }

        public static XAttribute GetOrCreateAttribute(this XElement el, XName name, string defaultValue = "")
        {
            if (el == null)
                throw new ArgumentNullException(nameof(el));

            var attr = el.Attribute(name);
            if (attr == null)
            {
                attr = new XAttribute(name, defaultValue);
                el.Add(attr);
            }

            return attr;
        }

        public static XElement FindByAttrVal(this IEnumerable<XElement> nodes, XName name, string attributeValue)
        {
            if (nodes == null)
                throw new ArgumentNullException(nameof(nodes));

            return nodes.FirstOrDefault(node => node.AttributeValue(name).Equals(attributeValue));
        }

        public static XAttribute GetValAttr(this XElement el)
        {
            if (el == null)
                return null;

            var valAttr = el.Attribute("val");
            return valAttr ?? el.Attribute(DocxNamespace.Main + "val");
        }

        public static string GetVal(this XElement el, string defaultValue = "")
        {
            return GetValAttr(el)?.Value ?? defaultValue;
        }

        public static string AttributeValue(this XElement el, XName name, string defaultValue = "")
        {
            var attr = el?.Attribute(name);
            return attr != null ? attr.Value : defaultValue;
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

                foreach (var item in xml.Descendants().Where(e => e.Name.LocalName == name))
                {
                    foreach (var child in LocalNameDescendants(item, localName))
                    {
                        yield return child;
                    }
                }
            }
            else
            {
                foreach (var item in xml.Descendants().Where(e => e.Name.LocalName == localName))
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
        internal static T GetEnumValue<T>(this XElement element)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(element));
            
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
        internal static T GetEnumValue<T>(this XAttribute attr)
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
        internal static void SetEnumValue<T>(this XElement element, T value)
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
        internal static bool TryGetEnumValue<T>(this XAttribute attr, out T result)
        {
            if (attr != null && !string.IsNullOrWhiteSpace(attr.Value))
            {
                string value = attr.Value;
                foreach (T e in Enum.GetValues(typeof(T)))
                {
                    FieldInfo fi = typeof(T).GetField(e.ToString());
                    string name = fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? e.ToString();
                    if (string.Compare(name, value, StringComparison.OrdinalIgnoreCase) == 0)
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
        internal static bool TryGetEnumValue<T>(this string value, out T result)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                foreach (T e in Enum.GetValues(typeof(T)))
                {
                    FieldInfo fi = typeof(T).GetField(e.ToString());
                    string name = fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? e.ToString();
                    if (string.Compare(name, value, StringComparison.OrdinalIgnoreCase) == 0)
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
        internal static bool TryGetEnumValue<T>(this XElement element, out T result)
        {
            return TryGetEnumValue(element?.GetVal(), out result);
        }

        /// <summary>
        /// Return xml string for this value
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        internal static string GetEnumName<T>(this T value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            FieldInfo fi = typeof(T).GetField(value.ToString());
            return fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? value.ToCamelCase();
        }
    }
}