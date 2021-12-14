using System;
using System.Drawing;
using System.Globalization;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Shapes;

namespace DXPlus
{
    /// <summary>
    /// Represents a drawing (vector or image) in this document.
    /// </summary>
    public class Picture : DocXElement
    {
        /// <summary>
        /// GUID for the decorative image extension
        /// </summary>
        private const string DecorativeImageId = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

        // English Metric Units are used for coordinates and sizes of drawings/pictures.
        // 1in == 914400 EMUs. Pixels are measured at 72dpi.
        private const int EmusInPixel = 914400 / 72;

        // The binary image being rendered by this Picture element. Note that images can be 
        // shared across different pictures.
        private readonly Image image;

        /// <summary>
        /// The non-visual properties for the drawing object. This should always be present.
        /// </summary>
        private XElement DocPr => Xml.Descendants(Namespace.WordProcessingDrawing + "docPr").Single();

        /// <summary>
        /// The non-visual properties for the picture contained in the drawing object.
        /// </summary>
        private XElement CNvPr => Xml.Descendants(Namespace.WordProcessingDrawing + "cNvPr").SingleOrDefault();

        /// <summary>
        /// Wraps a drawing or pict element in Word XML.
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">The XElement to wrap</param>
        /// <param name="image">The image to display</param>
        internal Picture(IDocument document, XElement xml, Image image) : base(document, xml)
        {
            if (xml.Name.LocalName != "drawing")
                throw new ArgumentException("Root element must be <drawing> for picture.");

            this.image = image;
        }

        private XElement spPr => Xml.FirstLocalNameDescendant("spPr");

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
        /// The unique id assigned to this drawing within the document.
        /// </summary>
        public int DrawingId => int.Parse(DocPr.AttributeValue("id"));

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
        /// Gets or sets the name of this Image.
        /// </summary>
        public string Name
        {
            get => DocPr.AttributeValue("name");
            set
            {
                DocPr.SetAttributeValue("name", value ?? "");
                CNvPr?.SetAttributeValue("name", value ?? "");
            }
        }

        /// <summary>
        /// Specifies whether this DrawingML object shall be displayed.
        /// </summary>
        public bool Hidden
        {
            get => DocPr.BoolAttributeValue("hidden");
            set
            {
                DocPr.SetAttributeValue("hidden", value ? "1" : null);
                CNvPr?.SetAttributeValue("hidden", value ? "1" : null);
            }
        }

        /// <summary>
        /// Gets or sets the description (alt-tag) for this picture.
        /// </summary>
        public string Description
        {
            get => DocPr.AttributeValue("descr");
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    value = null;

                DocPr.SetAttributeValue("descr", value);
                CNvPr?.SetAttributeValue("descr", value);
            }
        }

        /// <summary>
        /// Returns whether the picture is decorative.
        /// </summary>
        public bool IsDecorative
        {
            get => Xml.Descendants(Namespace.ADec + "decorative").FirstOrDefault()?.BoolAttributeValue("val") == true;
            set => SetDecorativeExtension(value);
        }

        /// <summary>
        /// Method to change the [adec:decorative] extension value.
        /// </summary>
        /// <param name="value">True/False to add/remove</param>
        private void SetDecorativeExtension(in bool value)
        {
            if (value)
            {
                FindDrawingExtension(DocPr, DecorativeImageId, true)
                    .GetOrAddElement(Namespace.ADec + "decorative")
                    .SetElementValue("val", "1");

                FindDrawingExtension(CNvPr, DecorativeImageId, true)?
                    .GetOrAddElement(Namespace.ADec + "decorative")
                    .SetElementValue("val", "1");
            }
            else
            {
                // Remove the extension.
                FindDrawingExtension(DocPr, DecorativeImageId, false)?.Remove();
                FindDrawingExtension(CNvPr, DecorativeImageId, false)?.Remove();
            }
        }

        /// <summary>
        /// Method to scan an extension list for a specific extension id
        /// </summary>
        /// <param name="owner">List owner</param>
        /// <param name="id">ID to look for</param>
        /// <param name="create">True to create it</param>
        /// <returns>XElement of [a:ext]</returns>
        private XElement FindDrawingExtension(XElement owner, string id, bool create)
        {
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("Value cannot be null or empty.", nameof(id));
            if (owner == null)
                return null;

            var extList = owner.Element(Namespace.DrawingMain + "extLst");
            if (extList == null && create)
            {
                extList = new XElement(Namespace.DrawingMain + "extLst");
                owner.Add(extList);
            }

            var extension = extList?.Elements(Namespace.DrawingMain + "ext")
                .SingleOrDefault(e => e.AttributeValue("uri") == id);
            if (extension == null && create)
            {
                extension = new XElement(Namespace.DrawingMain + "ext",
                    new XAttribute("uri", id));
            }

            return extension;
        }

        ///<summary>
        /// Returns the name of the image file for the picture.
        ///</summary>
        public string FileName => image.FileName;

        /// <summary>
        /// Get or sets the width of the rendered picture in pixels.
        /// </summary>
        public double Width
        {
            get
            {
                if (!double.TryParse(Xml.DescendantAttributeValues("cx").FirstOrDefault(), out var cx))
                {
                    var style = Xml.DescendantAttributes("style").FirstOrDefault();
                    if (style != null)
                    {
                        const string widthStr = "width:";
                        var fromWidth = style.Value.Substring(style.Value.IndexOf(widthStr, StringComparison.Ordinal) + widthStr.Length);
                        cx = double.Parse(fromWidth.Substring(0,
                                            fromWidth.IndexOf("pt", StringComparison.Ordinal)),
                                                CultureInfo.InvariantCulture) / EmusInPixel;
                    }
                }

                return cx / EmusInPixel;
            }

            set => Xml.DescendantAttributes("cx").ToList().ForEach(a => a.Value = (value * EmusInPixel).ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Get or sets the height of the rendered picture in pixels.
        /// </summary>
        public double Height
        {
            get
            {
                if (!double.TryParse(Xml.DescendantAttributeValues("cy").FirstOrDefault(), out var cy))
                {
                    var style = Xml.DescendantAttributes("style").FirstOrDefault();
                    if (style != null)
                    {
                        const string widthStr = "width:";
                        var fromWidth = style.Value.Substring(style.Value.IndexOf(widthStr, StringComparison.Ordinal) + widthStr.Length);
                        cy = double.Parse(fromWidth.Substring(0,
                                fromWidth.IndexOf("pt", StringComparison.Ordinal)),
                            CultureInfo.InvariantCulture) / EmusInPixel;
                    }
                }

                return cy / EmusInPixel;
            }

            set => Xml.DescendantAttributes("cy").ToList().ForEach(a => a.Value = (value * EmusInPixel).ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Remove this Picture from this document.
        /// </summary>
        public void Remove()
        {
            Xml.Remove();
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
    }
}