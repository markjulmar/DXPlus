using System;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This class is the base for text/breaks in a Run object.
    /// </summary>
    public class TextElement : ITextElement
    {
        protected readonly XElement Xml;

        /// <summary>
        /// Parent run object
        /// </summary>
        public Run Parent { get; }

        /// <summary>
        /// Name for this element.
        /// </summary>
        public string ElementType => Xml.Name.LocalName;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="runOwner"></param>
        /// <param name="xml"></param>
        internal TextElement(Run runOwner, XElement xml)
        {
            this.Parent = runOwner ?? throw new ArgumentNullException(nameof(runOwner));
            this.Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        }
    }
}