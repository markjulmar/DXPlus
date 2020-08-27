using System.Collections.Generic;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Container for identified sections in a document
    /// </summary>
    public class Section : Container
    {
        /// <summary>
        /// Create a new section
        /// </summary>
        /// <param name="document"></param>
        /// <param name="xml"></param>
        internal Section(IDocument document, XElement xml) : base(document, xml)
        {
        }

        /// <summary>
        /// Section break type (page)
        /// </summary>
        public SectionBreakType SectionBreakType { get; set; }

        /// <summary>
        /// Paragraphs in this section
        /// </summary>
        public List<Paragraph> SectionParagraphs { get; set; }
    }
}