using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Raw text from a Run. Wraps a {w:t} element.
/// </summary>
public class Text : TextElement
{
    static readonly XName xspace = XNamespace.Xml + "space";

    /// <summary>
    /// True if this text element preserves spaces.
    /// </summary>
    public bool PreserveSpaces
    {
        get => Xml.AttributeValue(xspace) == "preserve";
        set
        {
            if (!value)
            {
                Xml.Attribute(xspace)?.Remove();
            }
            else
            {
                Xml.SetAttributeValue(xspace, "preserve");
            }
        }
    }

    /// <summary>
    /// Text value for this container
    /// </summary>
    public string Value
    {
        get => Xml.Value;
        set => Xml.Value = value;
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="runOwner"></param>
    /// <param name="xml"></param>
    internal Text(Run runOwner, XElement xml) : base(runOwner, xml)
    {
    }
}