using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus.Charts;

/// <summary>
/// This element contains the 2-D pie series for this chart.
/// </summary>
public sealed class PieChart : Chart
{
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