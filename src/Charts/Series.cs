using System.Collections;
using System.Drawing;
using System.Xml.Linq;

namespace DXPlus.Charts;

/// <summary>
/// Represents a chart series
/// </summary>
public class Series
{
    private readonly XElement numCache;
    private readonly XElement strCache;

    /// <summary>
    /// Series xml element
    /// </summary>
    internal XElement Xml { get; }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="name">Name of the series</param>
    public Series(string name) : this(InitialXml(name))
    {
    }

    /// <summary>
    /// Constructor that loads a series from an XElement
    /// </summary>
    /// <param name="xml">XML representing the series</param>
    internal Series(XElement xml)
    {
        Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        strCache = Xml.Element(Namespace.Chart + "cat")?.Element(Namespace.Chart + "strRef")?.Element(Namespace.Chart + "strCache")
                   ?? throw new DocumentFormatException(nameof(strCache), "Chart missing series names.");
        numCache = Xml.Element(Namespace.Chart + "val")?.Element(Namespace.Chart + "numRef")?.Element(Namespace.Chart + "numCache")
                   ?? throw new DocumentFormatException(nameof(numCache), "Chart missing series values.");
    }

    /// <summary>
    /// Color for this series
    /// </summary>
    public Color? Color
    {
        get
        {
            var colorElement = Xml.Element(Namespace.Chart + "spPr");
            return colorElement == null
                ? null
                : System.Drawing.Color.FromArgb(
                    int.Parse(
                    colorElement.Element(Namespace.DrawingMain + "solidFill")!
                                  .Element(Namespace.DrawingMain + "srgbClr").GetVal()!,
                            System.Globalization.NumberStyles.HexNumber));
        }
        set
        {
            var colorElement = Xml.Element(Namespace.Chart + "spPr");
            colorElement?.Remove();
            if (value != null)
            {
                colorElement = new XElement(
                    Namespace.Chart + "spPr",
                    new XElement(Namespace.DrawingMain + "solidFill",
                        new XElement(Namespace.DrawingMain + "srgbClr",
                            new XAttribute("val", value.Value.ToHex()))));
                Xml.GetOrAddElement(Namespace.Chart + "tx").AddAfterSelf(colorElement);
            }
        }
    }

    /// <summary>
    /// Bind the series to a collection of names/values using reflection
    /// </summary>
    /// <param name="list">Collection to bind to</param>
    /// <param name="categoryPropertyName">property name</param>
    /// <param name="valuePropertyName">value name</param>
    public void Bind(ICollection list, string categoryPropertyName, string valuePropertyName)
    {
        strCache.RemoveAll();
        numCache.RemoveAll();

        var ptCount = new XElement(Namespace.Chart + "ptCount", new XAttribute("val", list.Count));
        var formatCode = new XElement(Namespace.Chart + "formatCode", "General");

        strCache.Add(ptCount);
        numCache.Add(formatCode);
        numCache.Add(ptCount);

        int index = 0;
        foreach (object item in list)
        {
            var pt = new XElement(Namespace.Chart + "pt",
                new XAttribute("idx", index),
                new XElement(Namespace.Chart + "v", item.GetType().GetProperty(categoryPropertyName)!.GetValue(item, null)));
            strCache.Add(pt);

            pt = new XElement(Namespace.Chart + "pt",
                new XAttribute("idx", index),
                new XElement(Namespace.Chart + "v", item.GetType().GetProperty(valuePropertyName)!.GetValue(item, null)));
            numCache.Add(pt);
            index++;
        }
    }

    /// <summary>
    /// Bind a series to two collections (categories and values)
    /// </summary>
    /// <param name="categories">Categories</param>
    /// <param name="values">Values</param>
    public void Bind(IList categories, IList values)
    {
        if (categories.Count != values.Count)
            throw new ArgumentException($"Passed {nameof(categories)} count must equal to {nameof(values)} count", nameof(categories));

        strCache.RemoveAll();
        numCache.RemoveAll();

        var ptCount = new XElement(Namespace.Chart + "ptCount", new XAttribute("val", categories.Count));
        var formatCode = new XElement(Namespace.Chart + "formatCode", "General");

        strCache.Add(ptCount);
        numCache.Add(formatCode);
        numCache.Add(ptCount);

        for (int index = 0; index < categories.Count; index++)
        {
            var category = categories[index];
            if (category == null)
                throw new ArgumentNullException(nameof(categories));

            var value = values[index];
            if (value == null)
                throw new ArgumentNullException(nameof(values));

            var pt = new XElement(Namespace.Chart + "pt",
                new XAttribute("idx", index),
                new XElement(Namespace.Chart + "v", category.ToString()));
            strCache.Add(pt);

            pt = new XElement(Namespace.Chart + "pt",
                new XAttribute("idx", index),
                new XElement(Namespace.Chart + "v", value.ToString()));
            numCache.Add(pt);
        }
    }

    /// <summary>
    /// Return the initial XML for the series.
    /// </summary>
    /// <param name="name">Name of the series</param>
    /// <returns>Create XElement</returns>
    private static XElement InitialXml(string name) =>
        XElement.Parse(
            $@"<c:ser xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
                        <c:tx>
                            <c:strRef>
                                <c:f></c:f>
                                <c:strCache>
                                    <c:pt idx=""0"">
                                        <c:v>{name}</c:v>
                                    </c:pt>
                                </c:strCache>
                            </c:strRef>
                        </c:tx>
                        <c:invertIfNegative>0</c:invertIfNegative>
                        <c:cat>
                            <c:strRef>
                                <c:f></c:f>
                                <c:strCache />
                            </c:strRef>
                        </c:cat>
                        <c:val>
                            <c:numRef>
                                <c:f></c:f>
                                <c:numCache />
                            </c:numRef>
                        </c:val>
                    </c:ser>");
}