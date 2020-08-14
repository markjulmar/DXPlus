using System.Collections.Generic;
using System.Xml.Linq;

namespace DXPlus
{
    public class Section : Container
    {
        internal Section(DocX document, XElement xml) : base(document, xml)
        {
        }

        public SectionBreakType SectionBreakType { get; set; }
        public List<Paragraph> SectionParagraphs { get; set; }
    }
}