using System;
using System.IO.Packaging;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Xml.Linq;

namespace DXPlus
{
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
        public int Id => int.Parse(DocPr.AttributeValue("id"));

        /// <summary>
        /// Parent run object (if any)
        /// </summary>
        public Run Parent => Xml.Parent?.Name == DXPlus.Name.Run 
            ? new Run(document, Xml.Parent, -1) : null;

        /// <summary>
        /// Element type (for ITextElement)
        /// </summary>
        public string ElementType => "drawing";
        
        /// <summary>
        /// Gets or sets the name of this drawing
        /// </summary>
        public string Name
        {
            get => DocPr.AttributeValue("name");
            set => DocPr.SetAttributeValue("name", value ?? "");
        }

        /// <summary>
        /// Specifies whether this DrawingML object shall be displayed.
        /// </summary>
        public bool Hidden
        {
            get => DocPr.BoolAttributeValue("hidden");
            set => DocPr.SetAttributeValue("hidden", value ? "1" : null);
        }
        
        /// <summary>
        /// Gets or sets the description (alt-tag) for this drawing
        /// </summary>
        public string Description
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
        /// Returns whether the picture is decorative.
        /// </summary>
        public bool IsDecorative
        {
            get => DrawingExtension.Get(DocPr, DrawingExtension.DecorativeImageId, false)?
                    .Element(Namespace.ADec + "decorative")?
                    .BoolAttributeValue("val") == true;
            set
            {
                if (value)
                {
                    DrawingExtension.Get(DocPr, DrawingExtension.DecorativeImageId, true)
                        .GetOrAddElement(Namespace.ADec + "decorative")
                        .SetElementValue("val", "1");
                }
                else
                {
                    // Remove the extension.
                    DrawingExtension.Get(DocPr, DrawingExtension.DecorativeImageId, false)?.Remove();
                }
            }
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
        internal Drawing(IDocument document, PackagePart packagePart, XElement xml) : base(document, packagePart, xml)
        {
            if (xml.Name.LocalName != "drawing")
                throw new ArgumentException("Root element must be <drawing> for Drawing.");
        }

        /// <summary>
        /// Retrieve the picture associated with this drawing.
        /// Null if it's not an image/picture.
        /// </summary>
        public Picture Picture
        {
            get
            {
                var pic = Xml.FirstLocalNameDescendant("pic"); // <pic:pic>
                var blip = pic?.FirstLocalNameDescendant("blip");
                if (blip == null) return null;
                
                // Retrieve the related image.
                var id = blip.AttributeValue(Namespace.RelatedDoc + "embed");
                if (!string.IsNullOrEmpty(id))
                    return new Picture(document, document.PackagePart, pic,
                        document.GetRelatedImage(id));
                
                // See if there's only an SVG. This happens on newer documents.
                var svgExtension = DrawingExtension.Get(blip, DrawingExtension.SvgExtensionId, false);
                if (svgExtension != null)
                {
                    id = svgExtension.FirstLocalNameDescendant("svgBlip")
                        .AttributeValue(Namespace.RelatedDoc + "embed");
                    if (!string.IsNullOrEmpty(id))
                        return new Picture(document, document.PackagePart, pic,
                            document.GetRelatedImage(id));
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
        public bool Equals(Drawing other)
        {
            if (other == null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Xml == other.Xml;
        }
   }
}