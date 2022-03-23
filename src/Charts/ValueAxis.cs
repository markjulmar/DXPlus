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
    internal ValueAxis(uint id) : base(InitialXml(id))
    {
    }

    /// <summary>
    /// Create the XML for a 
    /// </summary>
    /// <param name="id"></param>
    /// <returns></returns>
    private static XElement InitialXml(uint id) =>
        XElement.Parse($@"<c:valAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                <c:axId val=""{id}""/>
                <c:scaling>
                  <c:orientation val=""minMax""/>
                </c:scaling>
                <c:delete val=""0""/>
                <c:axPos val=""l""/>
                <c:numFmt sourceLinked=""0"" formatCode=""General""/>
                <c:majorGridlines/>
                <c:majorTickMark val=""out""/>
                <c:minorTickMark val=""none""/>
                <c:tickLblPos val=""nextTo""/>
                <c:crossAx val=""148921728""/>
                <c:crosses val=""autoZero""/>
                <c:crossBetween val=""between""/>
              </c:valAx>");
}