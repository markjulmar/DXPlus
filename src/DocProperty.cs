using System;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a field in the document.
    /// This field displays the value stored in a document or custom property.
    /// </summary>
    public class DocProperty : DocXElement, IEquatable<DocProperty>
    {
        /// <summary>
        /// Name of the property
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Value of the property
        /// </summary>
        public string Value { get; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="document"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        internal DocProperty(IDocument document, PackagePart packagePart, XElement name, XElement value) : base(document, packagePart, name)
        {
            var dpre = new Regex("DOCPROPERTY (?<name>.*) \\*");

            // Check for a simple field
            string instr = Xml.AttributeValue(DXPlus.Name.Instr, null)?.Trim();
            if (instr != null)
            {
                Name = instr.Contains("DOCPROPERTY")
                    ? dpre.Match(instr).Groups["name"].Value.Trim('"')
                    : instr.Substring(0, instr.IndexOf(' '));
                Value = Xml.Descendants().First(e => e.Name == DXPlus.Name.Text)?.Value;
            }
            // Complex field
            else
            {
                var instrText = Xml.Descendants(Namespace.Main + "instrText").Single();
                string text = instrText.Value.Trim();
                Name = text.Contains("DOCPROPERTY")
                    ? dpre.Match(text).Groups["name"].Value.Trim('"')
                    : text.Substring(0, text.IndexOf(' '));
                Value = new Run(document, value, 0).Text;
            }
        }

        /// <summary>
        /// Determines equality for doc properties
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(DocProperty other)
        {
            if (other == null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Xml == other.Xml;
        }
    }
}
