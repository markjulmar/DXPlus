namespace DXPlus
{
    /// <summary>
    /// Represents the headers in a section of the document
    /// </summary>
    public class Headers
    {
        /// <summary>
        /// Header on first page
        /// </summary>
        public Header First { get; set; }
        
        /// <summary>
        /// Header on even page
        /// </summary>
        public Header Even { get; set; }

        /// <summary>
        /// Header on odd pages
        /// </summary>
        public Header Odd { get; set; }
    }
}
