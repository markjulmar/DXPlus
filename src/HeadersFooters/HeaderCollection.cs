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
        /// <param name="documentOwner">Document</param>
        /// <param name="sectionOwner">Section owner</param>
        internal HeaderCollection(DocX documentOwner, Section sectionOwner)
            : base(documentOwner, sectionOwner, "hdr", Relations.Header, "header")
        {
        }
    }
}
