using System.Xml.Linq;

namespace DXPlus.Charts
{
    /// <summary>
    /// This element contains the 2-D pie series for this chart.
    /// </summary>
    public class PieChart : Chart
    {
        public override bool IsAxisExist => false;
        public override short MaxSeriesCount => 1;

        protected override XElement CreateChartXml()
        {
            return XElement.Parse(
                @"<c:pieChart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                  </c:pieChart>");
        }
    }
}
