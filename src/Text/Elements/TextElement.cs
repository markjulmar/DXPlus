using System;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This class is the base for text/breaks in a Run object.
    /// </summary>
    public class TextElement
    {
        protected readonly XElement Xml;
        private readonly Run runOwner;

        /// <summary>
        /// Parent run object
        /// </summary>
        public Run Parent => runOwner;

        /// <summary>
        /// Name for this element.
        /// </summary>
        public string Name => Xml.Name.LocalName;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="runOwner"></param>
        internal TextElement(Run runOwner, XElement xml)
        {
            this.runOwner = runOwner ?? throw new ArgumentNullException(nameof(runOwner));
            this.Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        }
    }
}