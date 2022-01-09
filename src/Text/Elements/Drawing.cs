﻿using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This wraps a {w:drawing} element in a Run object.
    /// </summary>
    public class Drawing : TextElement
    {
        public Picture Picture
        {
            get 
            {
                var document = (Document) Parent.Document;

                var id = Xml.FirstLocalNameDescendant("blip").AttributeValue(Namespace.RelatedDoc + "embed");
                return !string.IsNullOrEmpty(id)
                    ? new Picture(document, document.PackagePart, Xml,
                        document.GetRelatedImage(id))
                    : null;
            }
        }

        public Drawing(Run runOwner, XElement xml) : base(runOwner, xml)
        {
        }
    }
}
