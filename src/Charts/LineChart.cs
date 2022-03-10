using System.Xml.Linq;

namespace DXPlus.Charts;

/// <summary>
/// Specifies the kind of grouping for a column, line, or area chart.
/// </summary>
public enum Grouping
{
    /// <summary>
    /// Grouped by stacked percentage
    /// </summary>
    PercentStacked,
    /// <summary>
    /// Stacked
    /// </summary>
    Stacked,
    /// <summary>
    /// Standard grouping
    /// </summary>
    Standard
}

/// <summary>
/// This element contains the 2-D line chart series.
/// </summary>
public class LineChart : Chart
{
    /// <summary>
    /// Specifies the kind of grouping for a column, line, or area chart.
    /// </summary>
    public Grouping Grouping
    {
        get => ChartXml.Element(Namespace.Chart + "grouping")!.GetEnumValue<Grouping>();
        set => ChartXml.GetOrAddElement(Namespace.Chart + "grouping").SetEnumValue(value);
    }

    /// <summary>
    /// Create the initial chart XML
    /// </summary>
    protected override XElement CreateChartXml()
    {
        return XElement.Parse(
            @"<c:lineChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                 <c:grouping val=""standard""/>                    
              </c:lineChart>");
    }
}