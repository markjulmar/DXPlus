using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a field of type document property. This field displays the value stored in a custom property.
    /// </summary>
    public class DocProperty : DocXElement
    {
        /// <summary>
        /// Name of the property
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Value of the property
        /// </summary>
        public string Value { get; }

        internal DocProperty(DocX document, XElement xml)
            : base(document, xml)
        {
            string instr = Xml.AttributeValue(DocxNamespace.Main + "instr").Trim();
            Name = new Regex("DOCPROPERTY (?<name>.*) \\\\\\*").Match(instr).Groups["name"].Value;
            Value = Xml.Descendants().First(e => e.Name == DocxNamespace.Main + "t")?.Value;
        }
    }
}
