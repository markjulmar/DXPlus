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
    /// <param name="id">Identifier for the chart</param>
    internal CategoryAxis(uint id) : base(InitialXml(id))
    {
    }

    /// <summary>
    /// The initial XML used for a new chart axis.
    /// </summary>
    /// <param name="id"></param>
    /// <returns></returns>
    private static XElement InitialXml(uint id) =>
        XElement.Parse($@"<c:catAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                <c:axId val=""{id}""/>
                <c:scaling>
                  <c:orientation val=""minMax""/>
                </c:scaling>
                <c:delete val=""0""/>
                <c:axPos val=""b""/>
                <c:majorTickMark val=""out""/>
                <c:minorTickMark val=""none""/>
                <c:tickLblPos val=""nextTo""/>
                <c:crossAx val=""154227840""/>
                <c:crosses val=""autoZero""/>
                <c:auto val=""1""/>
                <c:lblAlgn val=""ctr""/>
                <c:lblOffset val=""100""/>
                <c:noMultiLvlLbl val=""0""/>
              </c:catAx>");
}