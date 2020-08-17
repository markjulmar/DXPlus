using System;
using System.Collections;
using System.Drawing;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus.Charts
{

    /// <summary>
    /// Represents a chart series
    /// </summary>
    public class Series
    {
        private XElement numCache;
        private XElement strCache;

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

        internal Series(XElement xml)
        {
            Xml = xml;
            LoadCache();
        }

        private void LoadCache()
        {
            strCache = Xml.Element(DocxNamespace.Chart + "cat").Element(DocxNamespace.Chart + "strRef").Element(DocxNamespace.Chart + "strCache");
            numCache = Xml.Element(DocxNamespace.Chart + "val").Element(DocxNamespace.Chart + "numRef").Element(DocxNamespace.Chart + "numCache");
        }

        public Color Color
        {
            get
            {
                XElement colorElement = Xml.Element(DocxNamespace.Chart + "spPr");
                return colorElement == null
                    ? Color.Transparent
                    : Color.FromArgb(int.Parse(
                        colorElement.Element(DocxNamespace.DrawingMain + "solidFill")
                                    .Element(DocxNamespace.DrawingMain + "srgbClr").GetVal(),
                        System.Globalization.NumberStyles.HexNumber));
            }
            set
            {
                XElement colorElement = Xml.Element(DocxNamespace.Chart + "spPr");
                colorElement?.Remove();
                colorElement = new XElement(
                    DocxNamespace.Chart + "spPr",
                    new XElement(DocxNamespace.DrawingMain + "solidFill",
                        new XElement(DocxNamespace.DrawingMain + "srgbClr",
                            new XAttribute("val", value.ToHex()))));
                Xml.Element(DocxNamespace.Chart + "tx").AddAfterSelf(colorElement);
            }
        }

        /// <summary>
        /// Series xml element
        /// </summary>
        internal XElement Xml { get; }

        public void Bind(ICollection list, string categoryPropertyName, string valuePropertyName)
        {
            strCache.RemoveAll();
            numCache.RemoveAll();

            XElement ptCount = new XElement(DocxNamespace.Chart + "ptCount", new XAttribute("val", list.Count));
            XElement formatCode = new XElement(DocxNamespace.Chart + "formatCode", "General");

            strCache.Add(ptCount);
            numCache.Add(formatCode);
            numCache.Add(ptCount);

            int index = 0;
            foreach (object item in list)
            {
                XElement pt = new XElement(DocxNamespace.Chart + "pt",
                            new XAttribute("idx", index),
                            new XElement(DocxNamespace.Chart + "v", item.GetType().GetProperty(categoryPropertyName).GetValue(item, null)));
                strCache.Add(pt);

                pt = new XElement(DocxNamespace.Chart + "pt",
                            new XAttribute("idx", index),
                            new XElement(DocxNamespace.Chart + "v", item.GetType().GetProperty(valuePropertyName).GetValue(item, null)));
                numCache.Add(pt);
                index++;
            }
        }

        public void Bind(IList categories, IList values)
        {
            if (categories.Count != values.Count)
            {
                throw new ArgumentException("Categories count must equal to Values count", nameof(categories));
            }

            strCache.RemoveAll();
            numCache.RemoveAll();

            XElement ptCount = new XElement(DocxNamespace.Chart + "ptCount", new XAttribute("val", categories.Count));
            XElement formatCode = new XElement(DocxNamespace.Chart + "formatCode", "General");

            strCache.Add(ptCount);
            numCache.Add(formatCode);
            numCache.Add(ptCount);

            for (int index = 0; index < categories.Count; index++)
            {
                XElement pt = new XElement(DocxNamespace.Chart + "pt",
                            new XAttribute("idx", index),
                            new XElement(DocxNamespace.Chart + "v", categories[index].ToString()));
                strCache.Add(pt);

                pt = new XElement(DocxNamespace.Chart + "pt",
                            new XAttribute("idx", index),
                            new XElement(DocxNamespace.Chart + "v", values[index].ToString()));
                numCache.Add(pt);
            }
        }
    }
}