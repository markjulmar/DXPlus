using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus.Tests
{
    public static class Helpers
    {
        public static XElement RemoveNamespaces(this XElement element)
        {
            if (!element.HasElements)
            {
                var xElement = new XElement(element.Name.LocalName) {Value = element.Value};
                foreach (var attribute in element.Attributes())
                    xElement.Add(new XAttribute(attribute.Name.LocalName, attribute.Value));
                return xElement;
            }

            return new XElement(element.Name.LocalName, element.Elements().Select(RemoveNamespaces));
        }

        public static IEnumerable<XElement> PathDescendants(this XContainer xml, string path)
        {
            if (xml == null)
                yield break;

            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentNullException(nameof(path));

            if (path.Contains('/'))
            {
                int pos = path.IndexOf('/');
                string name = path.Substring(0, pos);
                path = path.Substring(pos + 1);

                foreach (var item in xml.Descendants().Where(e => e.Name.LocalName == name))
                {
                    foreach (var child in PathDescendants(item, path))
                    {
                        yield return child;
                    }
                }
            }
            else
            {
                foreach (var item in xml.Descendants().Where(e => e.Name.LocalName == path))
                {
                    yield return item;
                }
            }
        }

    }
}
