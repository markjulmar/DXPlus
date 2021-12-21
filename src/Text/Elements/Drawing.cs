using System.Xml.Linq;

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
                if (!string.IsNullOrEmpty(id))
                {
                    var img = new Image(document, document?.PackagePart?.GetRelationship(id), id);
                    return new Picture(document, Xml, img) { PackagePart = document.PackagePart };
                }
                return null;
            }
        }

        public Drawing(Run runOwner, XElement xml) : base(runOwner, xml)
        {
        }
    }
}
