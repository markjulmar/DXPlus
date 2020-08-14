using System;
using System.Xml.Linq;
using System.Xml.Serialization;
using DXPlus.Helpers;

namespace DXPlus.Charts
{
    /// <summary>
    /// Specifies the possible directions for a bar chart.
    /// </summary>
    public enum BarDirection
    {
        [XmlAttribute("Col")]
        Column,
        Bar
    }

    /// <summary>
    /// Specifies the possible groupings for a bar chart.
    /// </summary>
    public enum BarGrouping
    {
        Clustered,
        PercentStacked,
        Stacked,
        Standard
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
            get => ChartXml.Element(DocxNamespace.Chart + "barDir").GetEnumValue<BarDirection>();
            set => ChartXml.Element(DocxNamespace.Chart + "barDir").SetEnumValue(value);
        }

        /// <summary>
        /// Specifies the possible groupings for a bar chart.
        /// </summary>
        public BarGrouping BarGrouping
        {
            get => ChartXml.Element(DocxNamespace.Chart + "grouping").GetEnumValue<BarGrouping>();
            set => ChartXml.Element(DocxNamespace.Chart + "grouping").SetEnumValue(value);
        }

        /// <summary>
        /// Specifies that its contents contain a percentage between 0% and 500%.
        /// </summary>
        public int GapWidth
        {
            get => Convert.ToInt32(ChartXml.Element(DocxNamespace.Chart + "gapWidth").GetVal());
            set
            {
                if ((value < 1) || (value > 500))
                {
                    throw new ArgumentException("GapWidth must be between 0% and 500%!", nameof(GapWidth));
                }

                ChartXml.Element(DocxNamespace.Chart + "gapWidth").Attribute("val").Value = value.ToString();
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
