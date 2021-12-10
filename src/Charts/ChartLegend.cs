using System.Xml.Linq;

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

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="position"></param>
        /// <param name="overlay"></param>
        internal ChartLegend(ChartLegendPosition position, bool overlay)
        {
            Xml = new XElement(Namespace.Chart + "legend",
                new XElement(Namespace.Chart + "legendPos", new XAttribute("val", position.GetEnumName())),
                new XElement(Namespace.Chart + "overlay", new XAttribute("val", overlay ? "1" : "0")));
        }

        /// <summary>
        /// Specifies that other chart elements shall be allowed to overlap this chart element
        /// </summary>
        public bool Overlay
        {
            get => Xml.Element(Namespace.Chart + "overlay").GetVal() == "1";
            set => Xml.GetOrAddElement(Namespace.Chart + "overlay").SetAttributeValue("val", value ? "1" : "0");
        }

        /// <summary>
        /// Specifies the possible positions for a legend
        /// </summary>
        public ChartLegendPosition Position
        {
            get => Xml.Element(Namespace.Chart + "legendPos").GetEnumValue<ChartLegendPosition>();
            set => Xml.GetOrAddElement(Namespace.Chart + "legendPos").SetEnumValue(value);
        }
    }
}