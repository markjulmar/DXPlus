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
        private static readonly Regex ExtractName = new Regex("DOCPROPERTY  (?<name>.*)  ");

        /// <summary>
        /// The custom property to display.
        /// </summary>
        public string Name { get; }

        internal DocProperty(DocX document, XElement xml)
            : base(document, xml)
        {
            string instr = Xml.AttributeValue(DocxNamespace.Main + "instr").Trim();
            Name = ExtractName.Match(instr).Groups["name"].Value;
        }
    }
}
