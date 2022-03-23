using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus.Charts;

/// <summary>
/// This element contains the 2-D bar or column series on this chart.
/// </summary>
public sealed class BarChart : Chart
{
    /// <summary>
    /// Default constructor for a BarChart.
    /// </summary>
    public BarChart() : base(Resources.Resource.BarChart(), true)
    {
    }

    /// <summary>
    /// Constructor for a loaded BarChart
    /// </summary>
    /// <param name="chartDocument"></param>
    internal BarChart(XDocument chartDocument) 
        : base(chartDocument, Namespace.Chart + "barChart")
    {
    }

    /// <summary>
    /// Specifies the possible directions for a bar chart.
    /// </summary>
    public BarDirection BarDirection
    {
        get => ChartXml.Element(Namespace.Chart + "barDir")!.GetEnumValue<BarDirection>();
        set => ChartXml.GetOrAddElement(Namespace.Chart + "barDir").SetEnumValue(value);
    }

    /// <summary>
    /// Specifies the possible groupings for a bar chart.
    /// </summary>
    public BarGrouping BarGrouping
    {
        get => ChartXml.Element(Namespace.Chart + "grouping")!.GetEnumValue<BarGrouping>();
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
            if (value is < 1 or > 500)
                throw new ArgumentException($"{nameof(GapWidth)} must be between 0 and 500.", nameof(GapWidth));

            ChartXml.GetOrAddElement(Namespace.Chart + "gapWidth").SetAttributeValue("val", value.ToString());
        }
    }
}