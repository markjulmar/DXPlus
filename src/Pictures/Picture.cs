using System;
using System.Drawing;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Shapes;

namespace DXPlus
{
    /// <summary>
    /// Represents a picture in this document.
    /// </summary>
    public sealed class Picture : DocXElement, IEquatable<Picture>
    {
        // The binary image being rendered by this Picture element. Note that images can be 
        // shared across different pictures.
        private readonly Image image;

        /// <summary>
        /// The non-visual properties for the picture contained in the drawing object.
        /// </summary>
        private XElement CNvPr => Xml.Descendants(Namespace.Picture + "cNvPr").SingleOrDefault();

        /// <summary>
        /// Wraps a drawing or pict element in Word XML.
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="packagePart">Package owner</param>
        /// <param name="xml">The XElement to wrap</param>
        /// <param name="image">The image to display</param>
        internal Picture(IDocument document, PackagePart packagePart, XElement xml, Image image) : base(document, packagePart, xml)
        {
            if (xml.Name.LocalName != "pic")
                throw new ArgumentException("Root element must be <pic> for picture.");

            this.image = image;
        }

        /// <summary>
        /// Shape properties section
        /// </summary>
        private XElement spPr => Xml.FirstLocalNameDescendant("spPr");

        /// <summary>
        /// Binary image properties
        /// </summary>
        private XElement blip => Xml.FirstLocalNameDescendant("blip");

        /// <summary>
        /// True if this image includes an embedded .svg. In this case,
        /// we always have a .png which is rendered in the document and a
        /// .svg which is _also_ in the relationship.
        /// </summary>
        public bool HasRelatedSvg => DrawingExtension.Get(blip, DrawingExtension.SvgExtensionId, false) != null;

        /// <summary>
        /// Specifies a flag indicating that the local BLIP compression
        /// setting overrides the document default compression setting.
        /// </summary>
        public bool? UseLocalDpi
        {
            get
            {
                var val = DrawingExtension.Get(blip, DrawingExtension.UseLocalDpiExtensionId, false)?
                        .FirstLocalNameDescendant("useLocalDpi")
                        .GetVal(null);
                if (val == null) return null;
                return (val == "1" || val.ToLower() == "true");
            }

            set =>
                DrawingExtension.Get(blip, DrawingExtension.UseLocalDpiExtensionId, true)
                    .GetOrAddElement(Namespace.DrawingA14 + "useLocalDpi")
                    .SetAttributeValue("val", value==true ? "1" : "0");
        }

        /// <summary>
        /// Returns the drawing owner of this picture.
        /// </summary>
        public Drawing Drawing
        {
            get
            {
                var xml = Xml;
                while (xml.Parent != null)
                {
                    if (xml.Parent.Name.LocalName == "drawing")
                    {
                        return new Drawing(document, packagePart, xml.Parent);
                    }
                }
                return null;
            }
        }
        
        /// <summary>
        /// Returns the relationship ID for the related SVG (if any).
        /// </summary>
        public string SvgRelationshipId
        {
            get => DrawingExtension.Get(blip, DrawingExtension.SvgExtensionId, false)?
                    .FirstLocalNameDescendant("svgBlip")
                    .AttributeValue(Namespace.RelatedDoc + "embed");

            set => DrawingExtension.Get(blip, DrawingExtension.SvgExtensionId, true)
                    .GetOrAddElement(Namespace.ASvg + "svgBlip")
                    .SetAttributeValue(Namespace.RelatedDoc + "embed", value);
        }

        /// <summary>
        /// Returns any related SVG image, null if none.
        /// </summary>
        public Image SvgImage
        {
            get
            {
                var rid = SvgRelationshipId;
                return !string.IsNullOrEmpty(rid) ? Document.GetRelatedImage(SvgRelationshipId) : null;
            }
        }

        /// <summary>
        /// Set a border line around the picture
        /// </summary>
        public Color? BorderColor
        {
            get
            {
                var shapeProperties = spPr;
                var border = shapeProperties?.Element(Namespace.DrawingMain + "ln");
                
                // Try solid color.
                var solidFill = border?.Element(Namespace.DrawingMain + "solidFill");
                if (solidFill != null)
                    return ShapeHelpers.ParseColorElement(solidFill.Element());

                return null;
            }

            set
            {
                var border = spPr.GetOrAddElement(Namespace.DrawingMain + "ln");
                border.Remove();
                if (value != null)
                {
                    spPr.GetOrAddElement(Namespace.DrawingMain + "ln")
                        .GetOrAddElement(Namespace.DrawingMain + "solidFill")
                        .GetOrAddElement(Namespace.DrawingMain + "srgbClr")
                        .SetAttributeValue("val", value.Value.ToHex());
                }
            }
        }

        /// <summary>
        /// Set a specific shape geometry for this picture
        /// </summary>
        /// <param name="shape">Shape</param>
        private void SetPictureShape(string shape)
        {
            Xml.FirstLocalNameDescendant("prstGeom")
               .SetAttributeValue("prst", shape);
        }

        /// <summary>
        /// Set the shape of this Picture to one in the BasicShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the BasicShapes enumeration.</param>
        public Picture SetPictureShape(BasicShapes shape)
        {
            SetPictureShape(shape.GetEnumName());
            return this;
        }

        /// <summary>
        /// Set the shape of this Picture to one in the RectangleShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the RectangleShapes enumeration.</param>
        public Picture SetPictureShape(RectangleShapes shape)
        {
            SetPictureShape(shape.GetEnumName());
            return this;
        }

        /// <summary>
        /// Set the shape of this Picture to one in the BlockArrowShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the BlockArrowShapes enumeration.</param>
        public Picture SetPictureShape(BlockArrowShapes shape)
        {
            SetPictureShape(shape.GetEnumName());
            return this;
        }

        /// <summary>
        /// Set the shape of this Picture to one in the EquationShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the EquationShapes enumeration.</param>
        public Picture SetPictureShape(EquationShapes shape)
        {
            SetPictureShape(shape.GetEnumName());
            return this;
        }

        /// <summary>
        /// Set the shape of this Picture to one in the FlowchartShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the FlowchartShapes enumeration.</param>
        public Picture SetPictureShape(FlowchartShapes shape)
        {
            SetPictureShape(shape.GetEnumName());
            return this;
        }

        /// <summary>
        /// Set the shape of this Picture to one in the StarAndBannerShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the StarAndBannerShapes enumeration.</param>
        public Picture SetPictureShape(StarAndBannerShapes shape)
        {
            SetPictureShape(shape.GetEnumName());
            return this;
        }

        /// <summary>
        /// Set the shape of this Picture to one in the CalloutShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the CalloutShapes enumeration.</param>
        public Picture SetPictureShape(CalloutShapes shape)
        {
            SetPictureShape(shape.GetEnumName());
            return this;
        }

        /// <summary>
        /// A unique id that identifies an Image embedded in this document.
        /// </summary>
        public string RelationshipId
        {
            get => Xml.FirstLocalNameDescendant("blip").AttributeValue(Namespace.RelatedDoc + "embed");
            set => Xml.FirstLocalNameDescendant("blip").SetAttributeValue(Namespace.RelatedDoc + "embed", value);
        }

        /// <summary>
        /// Flip this Picture Horizontally.
        /// </summary>
        public bool FlipHorizontal
        {
            get => Xml.FirstLocalNameDescendant("xfrm")?.BoolAttributeValue("flipH") == true;
            set => Xml.FirstLocalNameDescendant("xfrm").SetAttributeValue("flipH", value);
        }

        /// <summary>
        /// Flip this Picture Vertically.
        /// </summary>
        public bool FlipVertical
        {
            get => Xml.FirstLocalNameDescendant("xfrm")?.BoolAttributeValue("flipV") == true;
            set => Xml.FirstLocalNameDescendant("xfrm").SetAttributeValue("flipV", value);
        }

        /// <summary>
        /// The rotation in degrees of this image, actual value = value % 360
        /// </summary>
        public int Rotation
        {
            get => int.TryParse(Xml.FirstLocalNameDescendant("xfrm")?
                    .AttributeValue("rot", "0"), out var result)
                    ? (result / 60000) % 360
                    : 0;

            set
            {
                var rotation = (value % 360) * 60000;
                Xml.FirstLocalNameDescendant("xfrm")
                    .SetAttributeValue("rot", rotation);
            }
        }

        /// <summary>
        /// Id assigned to this picture
        /// </summary>
        public string Id
        {
            get => CNvPr.AttributeValue("id");
            set => CNvPr?.SetAttributeValue("id", value ?? "");
        }
        
        /// <summary>
        /// Gets or sets the name of this picture.
        /// </summary>
        public string Name
        {
            get => CNvPr.AttributeValue("name");
            set => CNvPr?.SetAttributeValue("name", value ?? "");
        }

        /// <summary>
        /// Specifies whether this DrawingML object shall be displayed.
        /// </summary>
        public bool Hidden
        {
            get => CNvPr.BoolAttributeValue("hidden");
            set => CNvPr?.SetAttributeValue("hidden", value ? "1" : null);
        }

        /// <summary>
        /// Gets or sets the description (alt-tag) for this picture.
        /// </summary>
        public string Description
        {
            get => CNvPr.AttributeValue("descr");
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    value = null;
                CNvPr?.SetAttributeValue("descr", value);
            }
        }

        /// <summary>
        /// Returns whether the picture is decorative.
        /// </summary>
        public bool IsDecorative
        {
            get => DrawingExtension.Get(CNvPr, DrawingExtension.DecorativeImageId, false)?
                .Element(Namespace.ADec + "decorative")?
                .BoolAttributeValue("val") == true;
            set
            {
                if (value)
                {
                    DrawingExtension.Get(CNvPr, DrawingExtension.DecorativeImageId, true)
                        .GetOrAddElement(Namespace.ADec + "decorative")
                        .SetElementValue("val", "1");
                }
                else
                {
                    // Remove the extension.
                    DrawingExtension.Get(CNvPr, DrawingExtension.DecorativeImageId, false)?.Remove();
                }
            }
        }

        ///<summary>
        /// Returns the name of the image file for the picture.
        ///</summary>
        public string FileName => image.FileName;

        /// <summary>
        /// Get the underlying image
        /// </summary>
        public Image Image => image;

        /// <summary>
        /// Get or sets the X-offset of the rendered picture in pixels.
        /// </summary>
        public double OffsetX
        {
            get => double.TryParse(spPr.Element(Namespace.DrawingMain + "xfrm")?
                    .Element(Namespace.DrawingMain + "off")
                    .AttributeValue("x"), out var result) ? result : 0;

            set => spPr.GetOrAddElement(Namespace.DrawingMain + "xfrm")
                    .GetOrAddElement(Namespace.DrawingMain + "off")
                    .SetAttributeValue("x");
        }

        /// <summary>
        /// Get or sets the Y-offset of the rendered picture in pixels.
        /// </summary>
        public double OffsetY
        {
            get => double.TryParse(spPr.Element(Namespace.DrawingMain + "xfrm")?
                .Element(Namespace.DrawingMain + "off")
                .AttributeValue("y"), out var result) ? result : 0;

            set => spPr.GetOrAddElement(Namespace.DrawingMain + "xfrm")
                .GetOrAddElement(Namespace.DrawingMain + "off")
                .SetAttributeValue("y");
        }

        /// <summary>
        /// Get or sets the width of the rendered picture in pixels.
        /// </summary>
        public double Width
        {
            get
            {
                var ext = spPr.Element(Namespace.DrawingMain + "xfrm")?
                    .Element(Namespace.DrawingMain + "ext");
                if (ext == null) return 0;
                return double.Parse(ext.AttributeValue("cx")) / Uom.EmuConversion;
            }

            set
            {
                if (value is < 0 or > Drawing.MaxExtentValue)
                    throw new ArgumentOutOfRangeException($"{nameof(Height)} cannot exceed {Drawing.MaxExtentValue}", nameof(Width));
                var ext = spPr.GetOrAddElement(Namespace.DrawingMain + "xfrm")
                    .GetOrAddElement(Namespace.DrawingMain + "ext");
                ext.SetAttributeValue("cx", value * Uom.EmuConversion);
            }
        }

        /// <summary>
        /// Get or sets the height of the rendered picture in pixels.
        /// </summary>
        public double Height
        {
            get
            {
                var ext = spPr.Element(Namespace.DrawingMain + "xfrm")?
                    .Element(Namespace.DrawingMain + "ext");
                if (ext == null) return 0;
                return double.Parse(ext.AttributeValue("cy")) / Uom.EmuConversion;
            }

            set
            {
                if (value is < 0 or > Drawing.MaxExtentValue)
                    throw new ArgumentOutOfRangeException($"{nameof(Height)} cannot exceed {Drawing.MaxExtentValue}", nameof(Width));
                var ext = spPr.GetOrAddElement(Namespace.DrawingMain + "xfrm")
                    .GetOrAddElement(Namespace.DrawingMain + "ext");
                ext.SetAttributeValue("cy", value * Uom.EmuConversion);
            }
        }

        /// <summary>
        /// Get or create a relationship link to a picture
        /// </summary>
        /// <returns>Relationship id</returns>
        internal string GetOrCreateRelationship()
        {
            string uri = image.PackageRelationship.TargetUri.OriginalString;

            return PackagePart.GetRelationshipsByType(Namespace.RelatedDoc.NamespaceName + "/image")
                       .Where(r => r.TargetUri.OriginalString == uri)
                       .Select(r => r.Id)
                       .SingleOrDefault() ??
                   PackagePart.CreateRelationship(image.PackageRelationship.TargetUri,
                       TargetMode.Internal, Namespace.RelatedDoc.NamespaceName + "/image").Id;
        }

        /// <summary>
        /// Determines equality for pictures
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(Picture other)
        {
            if (other == null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Xml == other.Xml;
        }
    }
}