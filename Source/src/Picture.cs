using System;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a Picture in this document, a Picture is a customized view of an Image.
    /// </summary>
    public class Picture : DocXElement
    {
        private const int EmusInPixel = 9525;
        internal Image img;
        private string name;
        private string descr;
        private int cx, cy;
        private uint rotation;
        private bool hFlip, vFlip;
        private readonly XElement xfrm;
        private readonly XElement prstGeom;

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
        /// <param name="document"></param>
        /// <param name="imageDefinition">The XElement i to wrap</param>
        /// <param name="image"></param>
        internal Picture(DocX document, XElement imageDefinition, Image image) : base(document, imageDefinition)
        {
            img = image;

            Id = Xml.LocalNameDescendants("blip")
                     .Select(e => e.AttributeValue(DocxNamespace.RelatedDoc + "embed"))
                     .SingleOrDefault();

            if (Id == null)
            {
                Id = Xml.Descendants("imagedata")
                        .Select(e => e.AttributeValue(DocxNamespace.RelatedDoc + "id"))
                        .SingleOrDefault();
            }

            name = Xml.DescendantAttributeValues("name").FirstOrDefault() ?? Xml.DescendantAttributeValues("title").FirstOrDefault();
            descr = Xml.DescendantAttributeValues("descr").FirstOrDefault();
            cx = Xml.DescendantAttributeValues("cx").Select(value => int.Parse(value)).FirstOrDefault();
            if (cx == 0)
            {
                XAttribute style = Xml.DescendantAttributes("style").FirstOrDefault();
                string fromWidth = style.Value.Substring(style.Value.IndexOf("width:") + 6);
                double widthInt = double.Parse(fromWidth.Substring(0, fromWidth.IndexOf("pt")).Replace(".", ",")) / 72.0 * 914400;
                cx = Convert.ToInt32(widthInt);
            }

            cy = Xml.DescendantAttributeValues("cy").Select(value => int.Parse(value)).FirstOrDefault();
            if (cy == 0)
            {
                XAttribute style = Xml.DescendantAttributes("style").FirstOrDefault();
                string fromHeight = style.Value.Substring(style.Value.IndexOf("height:") + 7);
                double heightInt = ((double.Parse((fromHeight.Substring(0, fromHeight.IndexOf("pt"))).Replace(".", ","))) / 72.0) * 914400;
                cy = Convert.ToInt32(heightInt);
            }

            xfrm = Xml.FirstLocalNameDescendant("xfrm");
            if (xfrm != null)
            {
                string val = xfrm.AttributeValue(DocxNamespace.RelatedDoc + "rot");
                rotation = string.IsNullOrEmpty(val)
                    ? 0
                    : uint.Parse(val);
            }

            prstGeom = Xml.FirstLocalNameDescendant("prstGeom");
        }

        private Picture SetPictureShape(string shape)
        {
            XAttribute prst = prstGeom.Attribute(DocxNamespace.RelatedDoc + "prst");
            if (prst == null)
            {
                prstGeom.Add(new XAttribute(DocxNamespace.RelatedDoc + "prst", "rectangle"));
            }

            prstGeom.SetAttributeValue(DocxNamespace.RelatedDoc + "prst", shape);

            return this;
        }

        /// <summary>
        /// Set the shape of this Picture to one in the BasicShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the BasicShapes enumeration.</param>
        public Picture SetPictureShape(BasicShapes shape)
        {
            return SetPictureShape(shape.GetEnumName());
        }

        /// <summary>
        /// Set the shape of this Picture to one in the RectangleShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the RectangleShapes enumeration.</param>
        public Picture SetPictureShape(RectangleShapes shape)
        {
            return SetPictureShape(shape.GetEnumName());
        }

        /// <summary>
        /// Set the shape of this Picture to one in the BlockArrowShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the BlockArrowShapes enumeration.</param>
        public Picture SetPictureShape(BlockArrowShapes shape)
        {
            return SetPictureShape(shape.GetEnumName());
        }

        /// <summary>
        /// Set the shape of this Picture to one in the EquationShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the EquationShapes enumeration.</param>
        public Picture SetPictureShape(EquationShapes shape)
        {
            return SetPictureShape(shape.GetEnumName());
        }

        /// <summary>
        /// Set the shape of this Picture to one in the FlowchartShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the FlowchartShapes enumeration.</param>
        public Picture SetPictureShape(FlowchartShapes shape)
        {
            return SetPictureShape(shape.GetEnumName());
        }

        /// <summary>
        /// Set the shape of this Picture to one in the StarAndBannerShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the StarAndBannerShapes enumeration.</param>
        public Picture SetPictureShape(StarAndBannerShapes shape)
        {
            return SetPictureShape(shape.GetEnumName());
        }

        /// <summary>
        /// Set the shape of this Picture to one in the CalloutShapes enumeration.
        /// </summary>
        /// <param name="shape">A shape from the CalloutShapes enumeration.</param>
        public Picture SetPictureShape(CalloutShapes shape)
        {
            return SetPictureShape(shape.GetEnumName());
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
                xfrm.SetAttributeValue(DocxNamespace.RelatedDoc + "flipH", hFlip ? 1 : 0);
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
                xfrm.SetAttributeValue(DocxNamespace.RelatedDoc + "flipV", vFlip ? 1 : 0);
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
                xfrm.SetAttributeValue(DocxNamespace.RelatedDoc + "rot", rotation);
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
                foreach (XAttribute a in Xml.DescendantAttributes(DocxNamespace.RelatedDoc + "name"))
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
            get => descr;

            set
            {
                descr = value;

                foreach (XAttribute a in Xml.DescendantAttributes(DocxNamespace.RelatedDoc + "descr"))
                {
                    a.Value = descr;
                }
            }
        }

        ///<summary>
        /// Returns the name of the image file for the picture.
        ///</summary>
        public string FileName => img.FileName;

        /// <summary>
        /// Get or sets the Width of this Image.
        /// </summary>
        public int Width
        {
            get => cx / EmusInPixel;

            set
            {
                cx = value;

                foreach (XAttribute a in Xml.DescendantAttributes(DocxNamespace.RelatedDoc + "cx"))
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

                foreach (XAttribute a in Xml.DescendantAttributes(DocxNamespace.RelatedDoc + "cy"))
                {
                    a.Value = (cy * EmusInPixel).ToString();
                }
            }
        }
    }
}