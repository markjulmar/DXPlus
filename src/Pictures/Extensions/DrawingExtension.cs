using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Class to represent a drawing extension.
/// </summary>
public class DrawingExtension
{
    internal readonly XElement Xml;

    /// <summary>
    /// Extension identifier
    /// </summary>
    public string UriId => Xml.AttributeValue("uri") ?? throw new DocumentFormatException(nameof(UriId));

    /// <summary>
    /// Constructor to create a new extension
    /// </summary>
    /// <param name="uriId">ID for the extension</param>
    protected DrawingExtension(string uriId)
        : this(new XElement(Namespace.DrawingMain + "ext", new XAttribute("uri", uriId)))
    {
    }
        
    /// <summary>
    /// Constructor
    /// </summary>
    protected internal DrawingExtension(XElement xml)
    {
        if (xml.Name.LocalName != "ext")
            throw new ArgumentException("Root of extension tag must be ext.");
            
        this.Xml = xml;
    }
}