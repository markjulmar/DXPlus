namespace DXPlus
{
    /// <summary>
    /// Specifies the type of the current section.
    /// </summary>
    public enum SectionBreakType
    {
        /// <summary>
        /// Specifies a next page section break, which begins the new section on the
        /// following page.
        /// </summary>
        NextPage,
        /// <summary>
        /// Specifies an even page section break, which begins the new section on the
        /// next even-numbered page, leaving the next odd page blank if necessary.
        /// </summary>
        EvenPage,
        /// <summary>
        /// Specifies an odd page section break, which begins the new section on the next
        /// odd-numbered page, leaving the next even page blank if necessary.
        /// </summary>
        OddPage,
        /// <summary>
        /// Specifies a continuous section break, which begin the new section on the
        /// following paragraph. This means that continuous section breaks might not
        /// specify certain page-level section properties, since they must be inherited
        /// from the following section. These breaks, however, can specify other section
        /// properties, such as line numbering and footnote/endnote settings.
        /// </summary>
        Continuous,
        /// <summary>
        /// Specifies a column section break, which begins the new section on the
        /// following column on the page.
        /// </summary>
        NextColumn
    }
}