﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DXPlus.Charts
{
    /// <summary>
    /// Specifies the possible positions for a legend.
    /// </summary>
    public enum ChartLegendPosition
    {
        [XmlAttribute("t")]
        Top,
        [XmlAttribute("b")]
        Bottom,
        [XmlAttribute("l")]
        Left,
        [XmlAttribute("r")]
        Right,
        [XmlAttribute("tr")]
        TopRight
    }

    /// <summary>
    /// Specifies the possible ways to display blanks.
    /// </summary>
    public enum DisplayBlanksAs
    {
        Gap,
        Span,
        Zero
    }

    /// <summary>
    /// Represents every Chart in this document.
    /// </summary>
    public abstract class Chart
    {
        /// <summary>
        /// Create an Chart for this document
        /// </summary>
        internal Chart()
        {
            Xml = XDocument.Parse
                (@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
                   <c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
                       <c:roundedCorners val=""0""/>
                       <c:chart>
                           <c:autoTitleDeleted val=""0""/>
                           <c:plotVisOnly val=""1""/>
                           <c:dispBlanksAs val=""gap""/>
                           <c:showDLblsOverMax val=""0""/>
                       </c:chart>
                   </c:chartSpace>");

            // Create a real chart xml in an inheritor
            ChartXml = CreateChartXml();

            // Create result plotarea element
            var plotAreaXml = new XElement(Namespace.Chart + "plotArea",
                                        new XElement(Namespace.Chart + "layout"),
                                            ChartXml);

            // Set labels
            var dLblsXml = XElement.Parse(
                @"<c:dLbls xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                    <c:showLegendKey val=""0""/>
                    <c:showVal val=""0""/>
                    <c:showCatName val=""0""/>
                    <c:showSerName val=""0""/>
                    <c:showPercent val=""0""/>
                    <c:showBubbleSize val=""0""/>
                    <c:showLeaderLines val=""1""/>
                </c:dLbls>");
            ChartXml.Add(dLblsXml);

            if (HasAxis)
            {
                CategoryAxis = new CategoryAxis("148921728");
                ValueAxis = new ValueAxis("154227840");

                var axIDcatXml = XElement.Parse($@"<c:axId val=""{CategoryAxis.Id}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>");
                var axIDvalXml = XElement.Parse($@"<c:axId val=""{ValueAxis.Id}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""/>");

                var insertPoint = ChartXml.Element(Namespace.Chart + "gapWidth");
                if (insertPoint != null)
                {
                    insertPoint.AddAfterSelf(axIDvalXml);
                    insertPoint.AddAfterSelf(axIDcatXml);
                }
                else
                {
                    ChartXml.Add(axIDcatXml);
                    ChartXml.Add(axIDvalXml);
                }

                plotAreaXml.Add(CategoryAxis.Xml);
                plotAreaXml.Add(ValueAxis.Xml);
            }

            ChartRootXml = Xml.Root.Element(Namespace.Chart + "chart");
            ChartRootXml.Element(Namespace.Chart + "autoTitleDeleted").AddAfterSelf(plotAreaXml);
        }

        /// <summary>
        /// The category axis
        /// </summary>
        public CategoryAxis CategoryAxis { get; }

        /// <summary>
        /// Specifies how blank cells shall be plotted on a chart
        /// </summary>
        public DisplayBlanksAs DisplayBlanksAs
        {
            get => ChartRootXml.Element(Namespace.Chart + "dispBlanksAs").GetEnumValue<DisplayBlanksAs>();
            set => ChartRootXml.Element(Namespace.Chart + "dispBlanksAs").SetEnumValue(value);
        }

        /// <summary>
        /// Chart has an axis?
        /// </summary>
        public virtual bool HasAxis => true;

        /// <summary>
        /// Chart's legend.
        /// If legend doesn't exist property is null.
        /// </summary>
        public ChartLegend Legend { get; private set; }

        /// <summary>
        /// Return maximum count of series
        /// </summary>
        public virtual short MaxSeriesCount => short.MaxValue;

        /// <summary>
        /// Chart's series
        /// </summary>
        public List<Series> Series
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
        /// Represents the values axis
        /// </summary>
        public ValueAxis ValueAxis { get; }

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
        /// The XML representation of this chart
        /// </summary>
        internal XDocument Xml { get; }

        /// <summary>
        /// The root XML node of the chart (c:chart)
        /// </summary>
        private XElement ChartRootXml { get; }

        /// <summary>
        /// The root XML node of the specific chart (bar, pie, etc.)
        /// </summary>
        protected XElement ChartXml { get; }

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
            ChartRootXml.GetOrCreateElement(Namespace.Chart + "plotArea").AddAfterSelf(Legend.Xml);
        }

        /// <summary>
        /// Remove the legend from the chart.
        /// </summary>
        public void RemoveLegend()
        {
            if (HasLegend)
            {
                Legend.Xml.Remove();
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

        /// <summary>
        /// An abstract method which creates the current chart xml
        /// </summary>
        protected abstract XElement CreateChartXml();
    }
}