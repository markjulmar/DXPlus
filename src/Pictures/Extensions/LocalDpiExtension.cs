using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// A UseLocalDpi element that specifies a flag indicating that the local
/// BLIP compression setting overrides the document default compression setting.
/// </summary>
public sealed class LocalDpiExtension : DrawingExtension
{
    /// <summary>
    /// GUID for this extension
    /// </summary>
    public static string ExtensionId => "{28A0092B-C50C-407E-A947-70E740481C1C}";

    /// <summary>
    /// Flag value.
    /// </summary>
    public bool Value
    {
        get => Xml.Element(Namespace.DrawingA14 + "useLocalDpi").BoolAttributeValue("val") == true;
        set => Xml.GetOrAddElement(Namespace.DrawingA14 + "useLocalDpi")
            .SetAttributeValue("val", value ? "1" : "0");
    }

    /// <summary>
    /// Constructor for a new override
    /// </summary>
    /// <param name="value">Value</param>
    public LocalDpiExtension(bool value) : base(ExtensionId)
    {
        Value = value;
    }
        
    /// <summary>
    /// Constructor for an existing fragment
    /// </summary>
    /// <param name="xml"></param>
    internal LocalDpiExtension(XElement xml) : base(xml)
    {
        if (xml.AttributeValue("uri") != UriId)
            throw new ArgumentException("Invalid extension tag for LocalDpiExtension.");
    }
}