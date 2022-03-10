using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// This is used to mark an image as 'decorative'.
/// </summary>
public sealed class DecorativeImageExtension : DrawingExtension
{
    /// <summary>
    /// GUID for this extension
    /// </summary>
    public static string ExtensionId => "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

    /// <summary>
    /// Flag value.
    /// </summary>
    public bool Value
    {
        get => Xml.Element(Namespace.ADec + "decorative").BoolAttributeValue("val") == true;
        set => Xml.GetOrAddElement(Namespace.ADec + "decorative")
            .SetAttributeValue("val", value ? "1" : "0");
    }
        
    /// <summary>
    /// Constructor for decorative images
    /// </summary>
    /// <param name="value">True if this is a decorative image</param>
    public DecorativeImageExtension(bool value) : base(ExtensionId)
    {
        Value = value;
    }
        
    /// <summary>
    /// Constructor from an existing XML fragment.
    /// </summary>
    /// <param name="xml"></param>
    internal DecorativeImageExtension(XElement xml) : base(xml)
    {
        if (xml.AttributeValue("uri") != UriId)
            throw new ArgumentException("Invalid extension tag for DecorativeImageExtension.");
    }
}