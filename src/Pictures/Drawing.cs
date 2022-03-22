using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a drawing (vector or image) in this document.
/// </summary>
public sealed class Drawing: DocXElement, IEquatable<Drawing>, ITextElement
{
    internal const double MaxExtentValue = 27273042316900.0 / Uom.EmuConversion;
        
    /// <summary>
    /// The non-visual properties for the drawing object. This should always be present.
    /// </summary>
    private XElement DocPr => Xml.Descendants(Namespace.WordProcessingDrawing + "docPr").Single();

    /// <summary>
    /// The unique id assigned to this drawing within the document.
    /// </summary>
    public int Id => int.TryParse(DocPr.AttributeValue("id"), out var result) ? result : throw new DocumentFormatException(nameof(Id));

    /// <summary>
    /// Parent run object (if any)
    /// </summary>
    public Run? Parent => Xml.Parent?.Name == Internal.Name.Run 
        ? new Run(Document, PackagePart, Xml.Parent) : null;

    /// <summary>
    /// Element type (for ITextElement)
    /// </summary>
    public string ElementType => RunTextType.Drawing;

    /// <summary>
    /// Text length for this element
    /// </summary>
    public int Length => 0;

    /// <summary>
    /// Gets or sets the name of this drawing
    /// </summary>
    public string? Name
    {
        get => DocPr.AttributeValue("name");
        set => DocPr.SetAttributeValue("name", value ?? "");
    }

    /// <summary>
    /// Specifies whether this DrawingML object shall be displayed.
    /// </summary>
    public bool? Hidden
    {
        get => DocPr.BoolAttributeValue("hidden");
        set => DocPr.SetAttributeValue("hidden", value == true ? "1" : null);
    }
        
    /// <summary>
    /// Gets or sets the description (alt-tag) for this drawing
    /// </summary>
    public string? Description
    {
        get => DocPr.AttributeValue("descr");
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                value = null;
            DocPr.SetAttributeValue("descr", value);
        }
    }

    /// <summary>
    /// ImageExtensions associated with this drawing
    /// </summary>
    public ExtensionWrapper Extensions => new(Document, DocPr);
        
    /// <summary>
    /// True if this is a decorative picture
    /// </summary>
    public bool? IsDecorative
    {
        get => Extensions.Get<DecorativeImageExtension>()?.Value;
        set
        {
            if (value == true)
            {
                Extensions.Add(new DecorativeImageExtension(true));
            }
            else
            {
                Extensions.Remove(DecorativeImageExtension.ExtensionId);
            }
        }
    }
        
    /// <summary>
    /// Hyperlink to navigate to if the image is clicked.
    /// </summary>
    public Uri? Hyperlink
    {
        get => DocumentHelpers.GetHlinkClick(DocPr, this.PackagePart);
        set => DocumentHelpers.SetHlinkClick(DocPr, this.PackagePart, value);
    }

    /// <summary>
    /// Get or sets the width of the drawing in pixels.
    /// </summary>
    public double Width
    {
        get => double.Parse(Xml.Element(Namespace.WordProcessingDrawing + "inline")?
            .Element(Namespace.WordProcessingDrawing + "extent")?
            .AttributeValue("cx")??"0") / Uom.EmuConversion;
        set
        {
            if (value is < 0 or > MaxExtentValue)
                throw new ArgumentOutOfRangeException($"{nameof(Width)} cannot exceed {MaxExtentValue}", nameof(Width));
            Xml.GetOrAddElement(Namespace.WordProcessingDrawing + "inline")
                .GetOrAddElement(Namespace.WordProcessingDrawing + "extent")
                .SetAttributeValue("cx", value * Uom.EmuConversion);
        }
    }

    /// <summary>
    /// Get or sets the height of the drawing in pixels.
    /// </summary>
    public double Height
    {
        get => double.Parse(Xml.Element(Namespace.WordProcessingDrawing + "inline")?
            .Element(Namespace.WordProcessingDrawing + "extent")?
            .AttributeValue("cy")??"0") / Uom.EmuConversion;
        set
        {
            if (value is < 0 or > MaxExtentValue)
                throw new ArgumentOutOfRangeException($"{nameof(Height)} cannot exceed {MaxExtentValue}", nameof(Width));

            Xml.GetOrAddElement(Namespace.WordProcessingDrawing + "inline")
                .GetOrAddElement(Namespace.WordProcessingDrawing + "extent")
                .SetAttributeValue("cy", value * Uom.EmuConversion);
        }
    }

    /// <summary>
    /// Wraps a drawing or pict element in Word XML.
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package owner</param>
    /// <param name="xml">The XElement to wrap</param>
    internal Drawing(Document document, PackagePart packagePart, XElement xml) : base(xml)
    {
        if (xml.Name.LocalName != RunTextType.Drawing)
            throw new ArgumentException("Root element must be <drawing> for Drawing.");
        SetOwner(document, packagePart, false);
    }

    /// <summary>
    /// Retrieve the picture associated with this drawing.
    /// Null if it's not an image/picture.
    /// </summary>
    public Picture? Picture
    {
        get
        {
            var pic = Xml.FirstLocalNameDescendant("pic"); // <pic:pic>
            var blip = pic?.FirstLocalNameDescendant("blip");
            if (pic == null || blip == null) return null;
                
            // Retrieve the related image.
            var id = blip.AttributeValue(Namespace.RelatedDoc + "embed");
            if (string.IsNullOrEmpty(id))
            {
                // See if there's only an SVG. This is malformed, but Word still renders it.
                var svgExtension = blip.Element(Namespace.DrawingMain + "extLst")?
                    .Elements(Namespace.DrawingMain + "ext")
                    .SingleOrDefault(e => e.AttributeValue("uri") == SvgExtension.ExtensionId);
                if (svgExtension != null)
                {
                    id = svgExtension.FirstLocalNameDescendant("svgBlip")
                        .AttributeValue(Namespace.RelatedDoc + "embed");
                }
            }

            if (!string.IsNullOrEmpty(id))
            {
                var image = Document.GetRelatedImage(id);
                if (image != null)
                {
                    return new Picture(Document, PackagePart, pic, image);
                }
            }

            
            return null;
        }
    }
        
    /// <summary>
    /// Remove this drawing from this document.
    /// </summary>
    public void Remove()
    {
        Xml.Remove();
    }

    /// <summary>
    /// Determines equality for Drawing objects
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Drawing? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a drawing
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Drawing);

    /// <summary>
    /// Returns hashcode for this drawing
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();
}