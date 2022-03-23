using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus.Charts;

/// <summary>
/// This element contains the 2-D pie series for this chart.
/// </summary>
public sealed class PieChart : Chart
{
    /// <summary>
    /// Default constructor for a LineChart.
    /// </summary>
    public PieChart()
    {
    }

    /// <summary>
    /// Constructor for a loaded BarChart
    /// </summary>
    /// <param name="chartDocument"></param>
    internal PieChart(XDocument chartDocument)
        : base(chartDocument, Namespace.Chart + "pieChart")
    {
    }

    /// <summary>
    /// Chart has an axis?
    /// </summary>
    public override bool HasAxis => false;

    /// <summary>
    /// Return maximum count of series
    /// </summary>
    public override short MaxSeriesCount => 1;

    /// <summary>
    /// Method which creates the current chart XML
    /// </summary>
    protected override XElement CreateChartXml() => new(Namespace.Chart + "pieChart");
}