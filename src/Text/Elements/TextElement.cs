using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// This class is the base for text/breaks in a Run object.
/// </summary>
public class TextElement : XElementWrapper, ITextElement
{
    /// <summary>
    /// XML for this element
    /// </summary>
    protected new XElement Xml => base.Xml!;

    /// <summary>
    /// Parent run object
    /// </summary>
    public Run Parent { get; }

    /// <summary>
    /// Name for this element.
    /// </summary>
    public string ElementType => Xml.Name.LocalName;

    /// <summary>
    /// Length of this element
    /// </summary>
    public int Length => DocumentHelpers.GetSize(Xml);

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="runOwner"></param>
    /// <param name="xml"></param>
    internal TextElement(Run runOwner, XElement xml)
    {
        this.Parent = runOwner ?? throw new ArgumentNullException(nameof(runOwner));
        base.Xml = xml ?? throw new ArgumentNullException(nameof(xml));
    }
}