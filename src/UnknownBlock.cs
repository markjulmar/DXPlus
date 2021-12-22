using System;
using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This wraps unknown blocks found when enumerating the document.
    /// </summary>
    public class UnknownBlock : Block, IEquatable<UnknownBlock>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="packagePart"></param>
        /// <param name="xml"></param>
        internal UnknownBlock(IDocument doc, PackagePart packagePart, XElement xml) : base(doc, packagePart, xml)
        {
        }

        /// <summary>
        /// Returns the name of this block in the document.
        /// </summary>
        public string Name => Xml.Name.LocalName;

        /// <summary>
        /// Determines equality for an unknown block
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(UnknownBlock other)
        {
            if (other == null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Xml == other.Xml;
        }
    }
}