namespace DXPlus
{
    /// <summary>
    /// Specifies how the spacing between lines as specified in the line attribute is calculated. 
    /// </summary>
    public enum LineRule
    {
        /// <summary>
        /// Specifies that the line spacing of the parent object shall be automatically determined by the size of its contents,
        /// with no predetermined minimum or maximum size. Lines are measured as 240th of a line.
        /// </summary>
        Auto,

        /// <summary>
        /// Specifies that the height of the line shall be at least the value specified, but may be expanded to fit its content as needed.
        /// Lines are measured as 240th of a pt.
        /// </summary>
        AtLeast,

        /// <summary>
        /// Specifies that the height of the line shall be exactly the value specified, regardless of the size of the contents of the contents.
        /// If the contents are too large for the specified height, then they shall be clipped as necessary. Lines are measured as 240th of a pt.
        /// </summary>
        Exactly
    }
}
