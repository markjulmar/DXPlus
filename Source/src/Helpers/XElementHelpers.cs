using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DXPlus
{
    internal static class XElementHelpers
    {
        public static XElement GetOrCreateElement(this XElement el, XName name, string defaultValue = "")
        {
            XElement node = el.Element(name);
            if (node == null)
            {
                el.SetElementValue(name, defaultValue);
                node = el.Element(name);
            }
            
            return node;
        }

        public static XElement FindByAttrVal(this IEnumerable<XElement> nodes, XName name, string attributeValue)
        {
            return nodes.FirstOrDefault(node => node.AttributeValue(name).Equals(attributeValue));
        }

        public static XAttribute GetValAttr(this XElement el)
        {
            return el?.Attribute(DocxNamespace.Main + "val");
        }

        public static string GetVal(this XElement el, string defaultValue = "")
        {
            return el?.AttributeValue(DocxNamespace.Main + "val", defaultValue);
        }

        public static void SetVal(this XElement el, object value)
        {
            if (el == null)
                throw new ArgumentNullException(nameof(el));
            if (value == null)
                value = "";

            el.SetAttributeValue(DocxNamespace.Main + "val", value.ToString());
        }

        public static string AttributeValue(this XElement el, XName name, string defaultValue = "")
        {
            XAttribute attr = el?.Attribute(name);
            return attr != null ? attr.Value : defaultValue;
        }

        public static IEnumerable<XElement> LocalNameElements(this XContainer e, string localName)
        {
            return e.Elements().Where(e => e.Name.LocalName.Equals(localName));
        }

        public static XElement FirstLocalNameDescendant(this XContainer e, string localName)
        {
            return e.LocalNameDescendants(localName).FirstOrDefault();
        }

        public static IEnumerable<XElement> LocalNameDescendants(this XContainer e, string localName)
        {
            if (e == null) 
                yield break;

            if (string.IsNullOrWhiteSpace(localName))
                throw new ArgumentNullException(nameof(localName));

            if (localName.Contains('/'))
            {
                int pos = localName.IndexOf('/');
                string name = localName.Substring(0, pos);
                localName = localName.Substring(pos + 1);

                foreach (var item in e.Descendants().Where(e => e.Name.LocalName == name))
                {
                    foreach (var child in LocalNameDescendants(item, localName))
                    {
                        yield return child;
                    }
                }
            }
            else
            {
                foreach (var item in e.Descendants().Where(e => e.Name.LocalName == localName))
                {
                    yield return item;
                }
            }
        }

        public static IEnumerable<XAttribute> DescendantAttributes(this XContainer e, XName attribName)
        {
            return e.Descendants().Attributes(attribName);
        }

        public static IEnumerable<string> DescendantAttributeValues(this XElement e, XName attribName)
        {
            return e.DescendantAttributes(attribName).Select(a => a.Value);
        }

        /// <summary>
        /// Get value from XElement and convert it to enum
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        internal static T GetEnumValue<T>(this XElement element)
        {
            if (!TryGetEnumValue(element, out T result))
                throw new ArgumentException($"{element.GetVal()} could not be matched to enum {typeof(T).Name}.");

            return result;
        }

        /// <summary>
        /// Get value from XElement and convert it to enum
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        internal static T GetEnumValue<T>(this XAttribute attr)
        {
            if (!TryGetEnumValue(attr, out T result))
                throw new ArgumentException($"{attr.Value} could not be matched to enum {typeof(T).Name}.");

            return result;
        }

        /// <summary>
        /// Convert value to xml string and set it into XElement
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        internal static void SetEnumValue<T>(this XElement element, T value)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(element));

            element.SetAttributeValue("val", GetEnumName(value));
        }

        internal static bool TryGetEnumValue<T>(this XAttribute attr, out T result)
        {
            if (attr == null || string.IsNullOrWhiteSpace(attr.Value))
                throw new ArgumentNullException(nameof(attr));

            string value = attr.Value;
            foreach (T e in Enum.GetValues(typeof(T)))
            {
                FieldInfo fi = typeof(T).GetField(e.ToString());
                string name = fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? e.ToString();
                if (string.Compare(name, value, true) == 0)
                {
                    result = e;
                    return true;
                }
            }

            result = default;
            return false;
        }

        internal static bool TryGetEnumValue<T>(this XElement element, out T result)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(element));

            string value = element.GetVal();
            foreach (T e in Enum.GetValues(typeof(T)))
            {
                FieldInfo fi = typeof(T).GetField(e.ToString());
                string name = fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? e.ToString();
                if (string.Compare(name, value, true) == 0)
                {
                    result = e;
                    return true;
                }
            }

            result = default;
            return false;
        }

        /// <summary>
        /// Return xml string for this value
        /// </summary>
        /// <typeparam name="T">Enum type</typeparam>
        internal static string GetEnumName<T>(this T value)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));

            FieldInfo fi = typeof(T).GetField(value.ToString());
            return fi.GetCustomAttribute<XmlAttributeAttribute>()?.AttributeName ?? value.ToCamelCase();
        }
    }
}