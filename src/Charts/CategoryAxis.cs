using System.Xml.Linq;

namespace DXPlus.Charts;

/// <summary>
/// Represents Category Axes
/// </summary>
public sealed class CategoryAxis : Axis
{
    /// <summary>
    /// Constructor when loading from an existing document
    /// </summary>
    /// <param name="xml"></param>
    internal CategoryAxis(XElement xml) : base(xml)
    {
    }

    /// <summary>
    /// Constructor when creating a new chart
    /// </summary>
    /// <param name="id">Identifier for the axis</param>
    /// <param name="vaxId">Identifier for the cross axis</param>
    internal CategoryAxis(uint id, uint vaxId) : base(Resources.Resource.ChartCategoryAxis(id, vaxId))
    {
    }
}