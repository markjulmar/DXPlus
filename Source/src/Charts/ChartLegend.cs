using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus.Charts
{

    /// <summary>
    /// Represents a chart legend
    /// https://docs.microsoft.com/dotnet/api/documentformat.openxml.drawing.charts.legend
    /// </summary>
    public class ChartLegend
    {
        /// <summary>
        /// Legend xml element
        /// </summary>
        internal XElement Xml { get; }

        internal ChartLegend(ChartLegendPosition position, bool overlay)
        {
            Xml = new XElement(DocxNamespace.Chart + "legend",
                new XElement(DocxNamespace.Chart + "legendPos", new XAttribute("val", position.GetEnumName())),
                new XElement(DocxNamespace.Chart + "overlay", new XAttribute("val", overlay ? "1" : "0")));
        }

        /// <summary>
        /// Specifies that other chart elements shall be allowed to overlap this chart element
        /// </summary>
        public bool Overlay
        {
            get => Xml.Element(DocxNamespace.Chart + "overlay").GetVal() == "1";
            set => Xml.Element(DocxNamespace.Chart + "overlay").Attribute("val").Value = value ? "1" : "0";
        }

        /// <summary>
        /// Specifies the possible positions for a legend
        /// </summary>
        public ChartLegendPosition Position
        {
            get => Xml.Element(DocxNamespace.Chart + "legendPos").GetEnumValue<ChartLegendPosition>();
            set => Xml.Element(DocxNamespace.Chart + "legendPos").SetEnumValue(value);
        }
    }
}