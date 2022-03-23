using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus.Charts;

/// <summary>
/// This class manages all the charts in the document.
/// Each chart is stored in a separate XML file
/// </summary>
internal sealed class ChartManager
{
    private readonly Document document;
    private readonly Dictionary<string, (PackagePart packagePart, XDocument document)> loadedCharts = new();

    /// <summary>
    /// Constructor for the chart manager.
    /// </summary>
    internal ChartManager(Document document)
    {
        this.document = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <summary>
    /// Locate the chart section in a chart document
    /// </summary>
    /// <param name="chartDocument"></param>
    /// <returns></returns>
    private static Chart? Create(XDocument chartDocument)
    {
        // Locate the XML fragment representing the chart type. It should be under plotArea.
        var plotArea = chartDocument.Descendants(Namespace.Chart + "plotArea").Single();

        foreach (var child in plotArea.Descendants())
        {
            if (child.Name == Namespace.Chart + "barChart")
                return new BarChart(chartDocument);
            if (child.Name == Namespace.Chart + "lineChart")
                return new LineChart(chartDocument);
            if (child.Name == Namespace.Chart + "pieChart")
                return new PieChart(chartDocument);
        }

        return null;
    }

    /// <summary>
    /// Returns a chart by the identifier
    /// </summary>
    /// <param name="relationId">Relationship id from the drawing</param>
    /// <returns>Chart</returns>
    internal Chart? Get(string relationId)
    {
        if (string.IsNullOrWhiteSpace(relationId))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(relationId));

        if (loadedCharts.TryGetValue(relationId, out var result))
        {
            return Create(result.document);
        }

        var relationship = document.PackagePart.GetRelationship(relationId);

        Uri chartUri = relationship.SourceUri;
        var chartPackagePart = document.Package.GetPart(chartUri);

        var chartDocument = chartPackagePart.Load();
        loadedCharts.Add(relationId, (chartPackagePart, chartDocument));

        return Create(chartDocument);
    }

    /// <summary>
    /// Saves all the loaded charts back to the document.
    /// </summary>
    internal void Save()
    {
        foreach (var (packagePart, chartDocument) in loadedCharts.Values)
            packagePart.Save(chartDocument);
    }

    /// <summary>
    /// Create a new paragraph, append it to the document and add the specified chart to it
    /// </summary>
    internal (string relationId, long chartId) CreateRelationship(Chart chart)
    {
        // Create a new chart part uri.
        string chartPartUriPath;
        int chartIndex = 0;

        do
        {
            chartIndex++;
            chartPartUriPath = $"/word/charts/chart{chartIndex}.xml";
        } while (document.Package.PartExists(new Uri(chartPartUriPath, UriKind.Relative)));

        // Create chart part.
        var chartPackagePart = document.Package.CreatePart(new Uri(chartPartUriPath, UriKind.Relative), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", CompressionOption.Normal);

        // Create a new chart relationship
        string id = GetNextRelationshipId();
        _ = document.PackagePart.CreateRelationship(chartPackagePart.Uri, TargetMode.Internal, $"{Namespace.RelatedDoc.NamespaceName}/chart", id);

        // Save a chart info the chartPackagePart
        chartPackagePart.Save(chart.Xml);

        loadedCharts.Add(id, (chartPackagePart, chart.Xml));

        return (id, document.GetNextDocumentId());
    }

    /// <summary>
    /// Create a new relationship id by locating the last one used.
    /// </summary>
    /// <returns></returns>
    private string GetNextRelationshipId()
    {
        // Last used id (0 if none)
        int id = document.PackagePart.GetRelationships()
            .Where(r => r.Id[..3].Equals("rId"))
            .Select(r => int.TryParse(r.Id[3..], out var result) ? result : 0)
            .DefaultIfEmpty()
            .Max();
        return $"rId{id + 1}";
    }

}