using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus.Charts;

/// <summary>
/// This element contains the 2-D line chart series.
/// </summary>
public sealed class LineChart : Chart
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