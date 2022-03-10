﻿using System.Xml.Linq;

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
        public int? Id => int.TryParse(Xml.AttributeValue(DXPlus.Name.Id), out var result) ? result : null;

        /// <summary>
        /// Retrieve the associated comment.
        /// </summary>
        public Comment? Comment
        {
            get
            {
                int? id = Id;
                return id == null ? null : Parent.Document.GetComment(id.Value);
            }
        }

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
