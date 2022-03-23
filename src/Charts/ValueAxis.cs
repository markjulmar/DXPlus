using System.Xml.Linq;

namespace DXPlus.Charts;

/// <summary>
/// Represents Values Axes
/// </summary>
public sealed class ValueAxis : Axis
{
    /// <summary>
    /// Constructor used when loading from document.
    /// </summary>
    /// <param name="xml"></param>
    public ValueAxis(XElement xml) : base(xml)
    {
    }

    /// <summary>
    /// Constructor for new chart value axis.
    /// </summary>
    /// <param name="id">Unique identifier</param>
    /// <param name="caxId">Category axis id</param>
    internal ValueAxis(uint id, uint caxId) : base(Resources.Resource.ChartValueAxis(id, caxId))
    {
    }
}