using System;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DXPlus.Charts
{
    /// <summary>
    /// Specifies the possible directions for a bar chart.
    /// </summary>
    public enum BarDirection
    {
        [XmlAttribute("col")]
        Column,
        Bar
    }

    /// <summary>
    /// Specifies the possible groupings for a bar chart.
    /// </summary>
    public enum BarGrouping
    {
        PercentStacked,
        Clustered,
        Standard,
        Stacked
    }

    /// <summary>
    /// This element contains the 2-D bar or column series on this chart.
    /// </summary>
    public class BarChart : Chart
    {
        /// <summary>
        /// Specifies the possible directions for a bar chart.
        /// </summary>
        public BarDirection BarDirection
        {
            get => ChartXml.Element(Namespace.Chart + "barDir").GetEnumValue<BarDirection>();
            set => ChartXml.GetOrAddElement(Namespace.Chart + "barDir").SetEnumValue(value);
        }

        /// <summary>
        /// Specifies the possible groupings for a bar chart.
        /// </summary>
        public BarGrouping BarGrouping
        {
            get => ChartXml.Element(Namespace.Chart + "grouping").GetEnumValue<BarGrouping>();
            set => ChartXml.GetOrAddElement(Namespace.Chart + "grouping").SetEnumValue(value);
        }

        /// <summary>
        /// Specifies that its contents contain a percentage between 0% and 500%.
        /// </summary>
        public int GapWidth
        {
            get => Convert.ToInt32(ChartXml.Element(Namespace.Chart + "gapWidth").GetVal());
            set
            {
                if (value < 1 || value > 500)
                    throw new ArgumentException($"{nameof(GapWidth)} must be between 0% and 500%!", nameof(GapWidth));

                ChartXml.GetOrAddElement(Namespace.Chart + "gapWidth").SetAttributeValue("val", value.ToString());
            }
        }

        protected override XElement CreateChartXml()
        {
            return XElement.Parse(
                @"<c:barChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:barDir val=""col""/>
                    <c:grouping val=""clustered""/>                    
                    <c:gapWidth val=""150""/>
                  </c:barChart>");
        }
    }
}
