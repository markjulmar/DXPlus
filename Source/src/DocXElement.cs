using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// All DocX types are derived from DocXElement.
    /// This class contains properties which every element of a DocX must contain.
    /// </summary>
    public abstract class DocXElement
    {
        /// <summary>
        /// The section in the Package this element is represented by
        /// </summary>
        internal PackagePart packagePart;

        /// <summary>
        /// This is the actual Xml that gives this element substance.
        /// </summary>
        internal XElement Xml;

        /// <summary>
        /// This is a reference to the DocX object that this element belongs to.
        /// Every DocX element is connected to a document.
        /// </summary>
        internal DocX Document;

        /// <summary>
        /// Store both the document and xml so that they can be accessed by derived types.
        /// </summary>
        /// <param name="document">The document that this element belongs to.</param>
        /// <param name="xml">The Xml that gives this element substance</param>
        internal DocXElement(DocX document, XElement xml)
        {
            this.Document = document;
            this.Xml = xml;
        }
    }
}