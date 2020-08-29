namespace DXPlus
{
    /// <summary>
    /// Underlying elements for all document bits
    /// </summary>
    public interface IDocumentBase
    {
        /// <summary>
        /// This is a reference to the document object that this element belongs to.
        /// Every DocX element is connected to a document.
        /// </summary>
        IDocument Owner { get; }
    }
}