using System.Xml.Linq;

namespace DXPlus.Charts
{

    /// <summary>
    /// Represents Values Axes
    /// </summary>
    public class ValueAxis : Axis
    {
        public ValueAxis(string id)
        {
            Xml = XElement.Parse($@"<c:valAx xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
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

        internal ValueAxis(XElement xml)
                    : base(xml)
        {
        }
    }
}