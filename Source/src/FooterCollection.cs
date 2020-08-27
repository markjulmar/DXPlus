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
        /// <param name="documentOwner"></param>
        internal FooterCollection(DocX documentOwner) : base(documentOwner, "ftr", Relations.Footer, "footer")
        {
        }
    }
}
