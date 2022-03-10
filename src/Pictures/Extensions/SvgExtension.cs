using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// The SVG extension
/// </summary>
public sealed class SvgExtension : DrawingExtension
{
    /// <summary>
    /// The document owner
    /// </summary>
    readonly Document document;

    /// <summary>
    /// GUID for this extension
    /// </summary>
    public static string ExtensionId => "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    /// <summary>
    /// Relationship ID for the related SVG
    /// </summary>
    public string? RelationshipId
    {
        get => Xml.FirstLocalNameDescendant("svgBlip").AttributeValue(Namespace.RelatedDoc + "embed");
        set => Xml.GetOrAddElement(Namespace.ASvg + "svgBlip")
            .SetAttributeValue(Namespace.RelatedDoc + "embed", value);
    }
        
    /// <summary>
    /// Returns any related SVG image, null if none.
    /// </summary>
    public Image? Image
    {
        get
        {
            var rid = RelationshipId;
            return !string.IsNullOrEmpty(rid) ? document.GetRelatedImage(rid) : null;
        }
    }

    /// <summary>
    /// Constructor for a new extension with relationship id
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="rid">Relationship id</param>
    /// <returns></returns>
    internal SvgExtension(Document document, string rid) : base(ExtensionId)
    {
        this.document = document ?? throw new ArgumentNullException(nameof(document));
        RelationshipId = rid ?? throw new ArgumentNullException(nameof(rid));
    }
        
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="xml">XML</param>
    internal SvgExtension(Document document, XElement xml) : base(xml)
    {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        if (xml.AttributeValue("uri") != UriId)
            throw new ArgumentException("Invalid extension tag for Svg.");

        this.document = document ?? throw new ArgumentNullException(nameof(document));
    }
}