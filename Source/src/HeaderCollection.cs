namespace DXPlus
{
    /// <summary>
    /// Represents the headers in a section of the document
    /// </summary>
    public sealed class HeaderCollection : HeaderOrFooterCollection<Header>
    {
        /// <summary>
        /// Constructor used to create the header collection
        /// </summary>
        /// <param name="documentOwner"></param>
        internal HeaderCollection(DocX documentOwner) : base(documentOwner, "hdr", Relations.Header, "header")
        {
        }
    }
}
