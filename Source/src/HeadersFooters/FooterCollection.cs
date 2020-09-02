namespace DXPlus
{
    /// <summary>
    /// Footers found in a document section
    /// </summary>
    public sealed class FooterCollection : HeaderOrFooterCollection<Footer>
    {
        /// <summary>
        /// Constructor used to create the footers collection
        /// </summary>
        /// <param name="documentOwner">Document</param>
        /// <param name="sectionOwner">Section owner</param>
        internal FooterCollection(DocX documentOwner, Section sectionOwner)
            : base(documentOwner, sectionOwner, "ftr", Relations.Footer, "footer")
        {
        }
    }
}
