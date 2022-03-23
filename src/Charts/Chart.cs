using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus.Charts
{
    /// <summary>
    /// Represents every Chart in this document.
    /// </summary>
    public abstract class Chart
    {
        /// <summary>
        /// The document related to this chart
        /// </summary>
        internal XDocument Xml { get; }

        /// <summary>
        /// The root XML node of the specific chart (bar, pie, etc.)
        /// </summary>
        protected XElement ChartXml { get; }

        /// <summary>
        /// The root XML node of the chart (c:chart)
        /// </summary>
        private readonly XElement chartRootXml;

        /// <summary>
        /// Loaded chart from an existing document
        /// </summary>
        /// <param name="document">Root document for the chart</param>
        /// <param name="chartType">XName for the chart element</param>
        internal Chart(XDocument document, XName chartType)
        {
            this.Xml = document;
            this.ChartXml = document.Descendants(chartType).Single();
            this.chartRootXml = Xml.Root!.Element(Namespace.Chart + "chart")!;

            // Get the axis values
            var categoryAxisXml = document.FirstLocalNameDescendant("catAx");
            var valueAxisXml = document.FirstLocalNameDescendant("valAx");
            if (categoryAxisXml != null && valueAxisXml != null)
            {
                CategoryAxis = new CategoryAxis(categoryAxisXml);
                ValueAxis = new ValueAxis(valueAxisXml);
            }

            // Create the legend if it exists.
            var legendXml = chartRootXml.FirstLocalNameDescendant("legend");
            if (legendXml != null)
            {
                Legend = new ChartLegend(legendXml);
            }
        }

        /// <summary>
        /// Create an Chart for this document
        /// </summary>
        internal Chart(XElement chartXml, bool hasAxis)
        {
            Xml = Resources.Resource.BaseChart();
            ChartXml = chartXml;
            chartRootXml = Xml.Root!.Element(Namespace.Chart + "chart")!;

            // Insert the chart
            var plotAreaXml = Xml.Descendants(Namespace.Chart + "plotArea").Single();
            plotAreaXml.Add(ChartXml);

            if (hasAxis)
            {
                uint cxId = (uint) new Random().Next();
                uint vxId = (uint) new Random().Next();

                CategoryAxis = new CategoryAxis(cxId, vxId);
                ValueAxis = new ValueAxis(vxId, cxId);

                var categoryAxisId = XElement.Parse($@"<c:axId val=""{CategoryAxis.Id}"" xmlns:c=""{Namespace.Chart.NamespaceName}""/>");
                var valueAxisId = XElement.Parse($@"<c:axId val=""{ValueAxis.Id}"" xmlns:c=""{Namespace.Chart.NamespaceName}""/>");

                var insertPoint = ChartXml.Element(Namespace.Chart + "gapWidth");
                if (insertPoint != null)
                {
                    insertPoint.AddAfterSelf(valueAxisId);
                    insertPoint.AddAfterSelf(categoryAxisId);
                }
                else
                {
                    ChartXml.Add(categoryAxisId);
                    ChartXml.Add(valueAxisId);
                }

                plotAreaXml.Add(CategoryAxis!.Xml);
                plotAreaXml.Add(ValueAxis.Xml);
            }
        }

        /// <summary>
        /// Specifies how blank cells shall be plotted on a chart
        /// </summary>
        public DisplayBlanksAs DisplayBlanksAs
        {
            get => chartRootXml.Element(Namespace.Chart + "dispBlanksAs")!.GetEnumValue<DisplayBlanksAs>();
            set => chartRootXml.Element(Namespace.Chart + "dispBlanksAs")!.SetEnumValue(value);
        }

        /// <summary>
        /// Chart has an axis?
        /// </summary>
        public bool HasAxis => CategoryAxis != null || ValueAxis != null;

        /// <summary>
        /// The values axis (can be null).
        /// </summary>
        public ValueAxis? ValueAxis { get; }

        /// <summary>
        /// The Category axis (can be null).
        /// </summary>
        public CategoryAxis? CategoryAxis { get; }

        /// <summary>
        /// Chart's legend.
        /// If legend doesn't exist property is null.
        /// </summary>
        public ChartLegend? Legend { get; private set; }

        /// <summary>
        /// Return maximum count of series
        /// </summary>
        public virtual short MaxSeriesCount => short.MaxValue;

        /// <summary>
        /// Chart's series
        /// </summary>
        public IList<Series> Series
        {
            get
            {
                var series = new List<Series>();
                int index = 1;
                foreach (var element in ChartXml.Elements(Namespace.Chart + "ser"))
                {
                    element.Add(new XElement(Namespace.Chart + "idx"), index++);
                    series.Add(new Series(element));
                }
                return series;
            }
        }

        /// <summary>
        /// Get or set 3D view for this chart
        /// </summary>
        public bool View3D
        {
            get => ChartXml.Name.LocalName.Contains("3D");
            set
            {
                if (value)
                {
                    if (!View3D)
                    {
                        string currentName = ChartXml.Name.LocalName;
                        ChartXml.Name = Namespace.Chart + currentName.Replace("Chart", "3DChart");
                    }
                }
                else if (View3D)
                {
                    string currentName = ChartXml.Name.LocalName;
                    ChartXml.Name = Namespace.Chart + currentName.Replace("3DChart", "Chart");
                }
            }
        }

        /// <summary>
        /// True if this chart has a legend
        /// </summary>
        public bool HasLegend => Legend != null;

        /// <summary>
        /// Add standard legend to the chart.
        /// </summary>
        public void AddLegend() => AddLegend(ChartLegendPosition.Right, false);

        /// <summary>
        /// Add a legend with parameters to the chart.
        /// </summary>
        public void AddLegend(ChartLegendPosition position, bool overlay)
        {
            if (Legend != null)
                RemoveLegend();

            Legend = new ChartLegend(position, overlay);
            chartRootXml.GetOrAddElement(Namespace.Chart + "plotArea").AddAfterSelf(Legend.Xml);
        }

        /// <summary>
        /// Remove the legend from the chart.
        /// </summary>
        public void RemoveLegend()
        {
            if (HasLegend)
            {
                Legend?.Xml.Remove();
                Legend = null;
            }
        }

        /// <summary>
        /// Add a new series to this chart
        /// </summary>
        public void AddSeries(Series series)
        {
            int serCount = ChartXml.Elements(Namespace.Chart + "ser").Count();
            if (serCount >= MaxSeriesCount)
            {
                throw new InvalidOperationException($"{serCount} > {MaxSeriesCount} series count for {GetType().Name}");
            }

            series.Xml.AddFirst(new XElement(Namespace.Chart + "order", new XAttribute("val", (serCount + 1).ToString())));
            series.Xml.AddFirst(new XElement(Namespace.Chart + "idx", new XAttribute("val", (serCount + 1).ToString())));

            ChartXml.Add(series.Xml);
        }
   }
}