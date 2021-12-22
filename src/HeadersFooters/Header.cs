using System;

namespace DXPlus
{
    /// <summary>
    /// Represents a single header (even, odd, first) in a document
    /// </summary>
    public sealed class Header : HeaderOrFooter, IEquatable<Header>
    {
        /// <summary>
        /// Determines equality for a header
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(Header other)
        {
            if (other == null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Xml == other.Xml;
        }
    }
}