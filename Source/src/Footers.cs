namespace DXPlus
{
    /// <summary>
    /// Footers found in a document section
    /// </summary>
    public class Footers
    {
        /// <summary>
        /// Footer on first page
        /// </summary>
        public Footer First { get; set; }

        /// <summary>
        /// Footer on even pages
        /// </summary>
        public Footer Even { get; set; }

        /// <summary>
        /// Footer on odd pages
        /// </summary>
        public Footer Odd { get; set; }
    }
}
