using System;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a Picture in this document, a Picture is a customized view of an Image.
    /// </summary>
    public class Picture : DocXBase
    {
        private const int EmusInPixel = 9525;
        internal Image Image;

        private string name;
        private string description;
        private int cx, cy;
        private uint rotation;
        private bool hFlip, vFlip;
        private readonly XElement transformElement;
        private readonly XElement geometryElement;

        /// <summary>
        /// Remove this Picture from this document.
        /// </summary>
        public void Remove()
        {
            Xml.Remove();
        }

        /// <summary>
        /// Wraps an XElement as an Image
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="imageDefinition">The XElement to wrap</param>
        /// <param name="image">The image to display</param>
        internal Picture(IDocument document, XElement imageDefinition, Image image) : base(document, imageDefinition)
        {
            Image = image;

            Id = Xml.LocalNameDescendants("blip")
                .Select(e => e.AttributeValue(Namespace.RelatedDoc + "embed"))
                .SingleOrDefault() ?? Xml.Descendants("imagedata")
                .Select(e => e.AttributeValue(Namespace.RelatedDoc + "id"))
                .SingleOrDefault();

            var style = Xml.DescendantAttributes("style").FirstOrDefault();

            name = Xml.DescendantAttributeValues("name").FirstOrDefault() 
                   ?? Xml.DescendantAttributeValues("title").FirstOrDefault();

            description = Xml.DescendantAttributeValues("descr").FirstOrDefault();

            cx = Xml.DescendantAttributeValues("cx").Select(int.Parse).FirstOrDefault();
            if (cx == 0 && style != null)
            {
                string fromWidth = style.Value.Substring(style.Value.IndexOf("width:", StringComparison.Ordinal) + 6);
                double widthInt = double.Parse(fromWidth.Substring(0, fromWidth.IndexOf("pt", StringComparison.Ordinal)).Replace(".", ",")) / 72.0 * 914400;
                cx = Convert.ToInt32(widthInt);
            }

            cy = Xml.DescendantAttributeValues("cy").Select(int.Parse).FirstOrDefault();
            if (cy == 0 && style != null)
            {
                string fromHeight = style.Value.Substring(style.Value.IndexOf("height:", StringComparison.Ordinal) + 7);
                double heightInt = double.Parse((fromHeight.Substring(0, fromHeight.IndexOf("pt", StringComparison.Ordinal))).Replace(".", ",")) / 72.0 * 914400;
                cy = Convert.ToInt32(heightInt);
            }

            transformElement = Xml.FirstLocalNameDescendant("xfrm");
            if (transformElement != null)
            {
                string val = transformElement.AttributeValue(Namespace.RelatedDoc + "rot");
                rotation = string.IsNullOrEmpty(val)
                    ? 0
                    : uint.Parse(val);
            }

            geometryElement = Xml.FirstLocalNameDescendant("prstGeom");
        }

        /// <summary>
        /// Set a specific shape geometry for this picture
        /// </summary>
        /// <param name="shape">Shape</param>
        private void SetPictureShape(string shape)
        {
            var prst = geometryElement.Attribute(Namespace.RelatedDoc + "prst");
            if (prst == null)
                geometryElement.Add(new XAttribute(Namespace.RelatedDoc + "prst", "rectangle"));

            geometryElement.SetAttributeValue(Namespace.RelatedDoc + "prst", shape);
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
        public string Id { get; }

        /// <summary>
        /// Flip this Picture Horizontally.
        /// </summary>
        public bool FlipHorizontal
        {
            get => hFlip;

            set
            {
                hFlip = value;
                transformElement.SetAttributeValue(Namespace.RelatedDoc + "flipH", hFlip ? 1 : 0);
            }
        }

        /// <summary>
        /// Flip this Picture Vertically.
        /// </summary>
        public bool FlipVertical
        {
            get => vFlip;

            set
            {
                vFlip = value;
                transformElement.SetAttributeValue(Namespace.RelatedDoc + "flipV", vFlip ? 1 : 0);
            }
        }

        /// <summary>
        /// The rotation in degrees of this image, actual value = value % 360
        /// </summary>
        public uint Rotation
        {
            get => rotation / 60000;

            set
            {
                rotation = (value % 360) * 60000;
                transformElement.SetAttributeValue(Namespace.RelatedDoc + "rot", rotation);
            }
        }

        /// <summary>
        /// Gets or sets the name of this Image.
        /// </summary>
        public string Name
        {
            get => name;

            set
            {
                name = value;
                foreach (var a in Xml.DescendantAttributes(Namespace.RelatedDoc + "name"))
                {
                    a.Value = name;
                }
            }
        }

        /// <summary>
        /// Gets or sets the description for this Image.
        /// </summary>
        public string Description
        {
            get => description;

            set
            {
                description = value;

                foreach (var a in Xml.DescendantAttributes(Namespace.RelatedDoc + "descr"))
                {
                    a.Value = description;
                }
            }
        }

        ///<summary>
        /// Returns the name of the image file for the picture.
        ///</summary>
        public string FileName => Image.FileName;

        /// <summary>
        /// Get or sets the Width of this Image.
        /// </summary>
        public int Width
        {
            get => cx / EmusInPixel;

            set
            {
                cx = value;

                foreach (var a in Xml.DescendantAttributes(Namespace.RelatedDoc + "cx"))
                {
                    a.Value = (cx * EmusInPixel).ToString();
                }
            }
        }

        /// <summary>
        /// Get or sets the height of this Image.
        /// </summary>
        public int Height
        {
            get => cy / EmusInPixel;

            set
            {
                cy = value;

                foreach (var a in Xml.DescendantAttributes(Namespace.RelatedDoc + "cy"))
                {
                    a.Value = (cy * EmusInPixel).ToString();
                }
            }
        }
    }
}