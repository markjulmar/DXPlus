using System;
using System.Globalization;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a drawing (vector or image) in this document.
    /// </summary>
    public class Picture : DocXElement
    {
        // English Metric Units are used for coordinates and sizes of drawings/pictures.
        // 1in == 914400 EMUs. Pixels are measured at 72dpi.
        private const int EmusInPixel = 914400 / 72;

        // The binary image being rendered by this Picture element. Note that images can be 
        // shared across different pictures.
        internal Image Image;

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

            Image = image;
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
        public string Id
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
        /// Gets or sets the name of this Image.
        /// </summary>
        public string Name
        {
            get => Xml.DescendantAttributeValues("name").FirstOrDefault();
            set => Xml.DescendantAttributes("name").ToList().ForEach(a => a.Value = value ?? string.Empty);
        }

        /// <summary>
        /// Gets or sets the description for this Image.
        /// </summary>
        public string Description
        {
            get => Xml.DescendantAttributeValues("descr").FirstOrDefault();
            set => Xml.DescendantAttributes("descr").ToList().ForEach(a => a.Value = value ?? string.Empty);
        }

        ///<summary>
        /// Returns the name of the image file for the picture.
        ///</summary>
        public string FileName => Image.FileName;

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
            string uri = Image.PackageRelationship.TargetUri.OriginalString;

            return PackagePart.GetRelationshipsByType(Namespace.RelatedDoc.NamespaceName + "/image")
                       .Where(r => r.TargetUri.OriginalString == uri)
                       .Select(r => r.Id)
                       .SingleOrDefault() ??
                   PackagePart.CreateRelationship(Image.PackageRelationship.TargetUri,
                       TargetMode.Internal, Namespace.RelatedDoc.NamespaceName + "/image").Id;
        }
    }
}