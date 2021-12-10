using System;
using System.Collections;
using System.Drawing;
using System.Xml.Linq;

namespace DXPlus.Charts
{

    /// <summary>
    /// Represents a chart series
    /// </summary>
    public class Series
    {
        private XElement numCache;
        private XElement strCache;

        /// <summary>
        /// Series xml element
        /// </summary>
        internal XElement Xml { get; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="name">Name of the series</param>
        public Series(string name)
        {
            Xml = XElement.Parse(
                 $@"<c:ser xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                        <c:tx>
                            <c:strRef>
                                <c:f></c:f>
                                <c:strCache>
                                    <c:pt idx=""0"">
                                        <c:v>{name}</c:v>
                                    </c:pt>
                                </c:strCache>
                            </c:strRef>
                        </c:tx>
                        <c:invertIfNegative>0</c:invertIfNegative>
                        <c:cat>
                            <c:strRef>
                                <c:f></c:f>
                                <c:strCache />
                            </c:strRef>
                        </c:cat>
                        <c:val>
                            <c:numRef>
                                <c:f></c:f>
                                <c:numCache />
                            </c:numRef>
                        </c:val>
                    </c:ser>");

            LoadCache();
        }

        /// <summary>
        /// Constructor that loads a series from an XElement
        /// </summary>
        /// <param name="xml"></param>
        internal Series(XElement xml)
        {
            Xml = xml;
            LoadCache();
        }

        /// <summary>
        /// Load the cached XElement representation for the series names and values
        /// </summary>
        private void LoadCache()
        {
            strCache = Xml.Element(Namespace.Chart + "cat")?.Element(Namespace.Chart + "strRef")?.Element(Namespace.Chart + "strCache");
            numCache = Xml.Element(Namespace.Chart + "val")?.Element(Namespace.Chart + "numRef")?.Element(Namespace.Chart + "numCache");
        }

        /// <summary>
        /// Color for this series
        /// </summary>
        public Color Color
        {
            get
            {
                var colorElement = Xml.Element(Namespace.Chart + "spPr");
                return colorElement == null
                    ? Color.Transparent
                    : Color.FromArgb(int.Parse(
                        colorElement.Element(Namespace.DrawingMain + "solidFill")
                                    .Element(Namespace.DrawingMain + "srgbClr").GetVal(),
                        System.Globalization.NumberStyles.HexNumber));
            }
            set
            {
                var colorElement = Xml.Element(Namespace.Chart + "spPr");
                colorElement?.Remove();
                
                colorElement = new XElement(
                    Namespace.Chart + "spPr",
                    new XElement(Namespace.DrawingMain + "solidFill",
                        new XElement(Namespace.DrawingMain + "srgbClr",
                            new XAttribute("val", value.ToHex()))));
                Xml.GetOrAddElement(Namespace.Chart + "tx").AddAfterSelf(colorElement);
            }
        }

        /// <summary>
        /// Bind the series to a collection of names/values using reflection
        /// </summary>
        /// <param name="list">Collection to bind to</param>
        /// <param name="categoryPropertyName">property name</param>
        /// <param name="valuePropertyName">value name</param>
        public void Bind(ICollection list, string categoryPropertyName, string valuePropertyName)
        {
            strCache.RemoveAll();
            numCache.RemoveAll();

            var ptCount = new XElement(Namespace.Chart + "ptCount", new XAttribute("val", list.Count));
            var formatCode = new XElement(Namespace.Chart + "formatCode", "General");

            strCache.Add(ptCount);
            numCache.Add(formatCode);
            numCache.Add(ptCount);

            int index = 0;
            foreach (object item in list)
            {
                var pt = new XElement(Namespace.Chart + "pt",
                            new XAttribute("idx", index),
                            new XElement(Namespace.Chart + "v", item.GetType().GetProperty(categoryPropertyName).GetValue(item, null)));
                strCache.Add(pt);

                pt = new XElement(Namespace.Chart + "pt",
                            new XAttribute("idx", index),
                            new XElement(Namespace.Chart + "v", item.GetType().GetProperty(valuePropertyName).GetValue(item, null)));
                numCache.Add(pt);
                index++;
            }
        }

        /// <summary>
        /// Bind a series to two collections (categories and values)
        /// </summary>
        /// <param name="categories">Categories</param>
        /// <param name="values">Values</param>
        public void Bind(IList categories, IList values)
        {
            if (categories.Count != values.Count)
                throw new ArgumentException($"Passed {nameof(categories)} count must equal to {nameof(values)} count", nameof(categories));

            strCache.RemoveAll();
            numCache.RemoveAll();

            var ptCount = new XElement(Namespace.Chart + "ptCount", new XAttribute("val", categories.Count));
            var formatCode = new XElement(Namespace.Chart + "formatCode", "General");

            strCache.Add(ptCount);
            numCache.Add(formatCode);
            numCache.Add(ptCount);

            for (int index = 0; index < categories.Count; index++)
            {
                var pt = new XElement(Namespace.Chart + "pt",
                            new XAttribute("idx", index),
                            new XElement(Namespace.Chart + "v", categories[index].ToString()));
                strCache.Add(pt);

                pt = new XElement(Namespace.Chart + "pt",
                            new XAttribute("idx", index),
                            new XElement(Namespace.Chart + "v", values[index].ToString()));
                numCache.Add(pt);
            }
        }
    }
}