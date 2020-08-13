using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace DXPlus
{
    /// <summary>
    /// Represents a field of type document property. This field displays the value stored in a custom property.
    /// </summary>
    public class DocProperty: DocXElement
    {
        private static readonly Regex extractName = new Regex("DOCPROPERTY  (?<name>.*)  ");

        /// <summary>
        /// The custom property to display.
        /// </summary>
        public string Name { get; }

        internal DocProperty(DocX document, XElement xml)
            : base(document, xml)
        {
            string instr = Xml.AttributeValue(DocxNamespace.Main + "instr");
            Name = extractName.Match(instr.Trim()).Groups["name"].Value;
        }
    }
}
