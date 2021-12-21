using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This identifies a comment reference tied to a Run.
    /// </summary>
    public class CommentRef : TextElement
    {
        /// <summary>
        /// The comment identifier.
        /// </summary>
        public int Id => int.Parse(Xml.AttributeValue(DXPlus.Name.Id));

        /// <summary>
        /// Retrieve the associated comment.
        /// </summary>
        public Comment Comment => Parent.Document.GetComment(Id);

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="runOwner"></param>
        /// <param name="xml"></param>
        internal CommentRef(Run runOwner, XElement xml) : base(runOwner, xml)
        {
        }
    }
}
