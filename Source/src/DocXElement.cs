using System.ComponentModel;
using System.IO.Packaging;
using System.Runtime.CompilerServices;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// All DocX types are derived from DocXElement.
    /// This class contains properties which every element of a DocX must contain.
    /// </summary>
    public abstract class DocXElement
    {
        private DocX document;
        private XElement xml;

        /// <summary>
        /// The file in the package where this element is stored
        /// </summary>
        internal PackagePart PackagePart { get; set; }

        /// <summary>
        /// This is the actual Xml that gives this element substance.
        /// </summary>
        public XElement Xml
        {
            get => xml;
            set
            {
                var previousValue = xml;
                xml = value;
                OnElementChanged(previousValue, xml);
            }
        }

        /// <summary>
        /// Called when the XML element is changed
        /// </summary>
        protected virtual void OnElementChanged(XElement previousValue, XElement newValue)
        {
        }

        /// <summary>
        /// This is a reference to the document object that this element belongs to.
        /// Every DocX element is connected to a document.
        /// </summary>
        public DocX Document
        {
            get => document;
            set
            {
                var previousValue = document;
                document = value;
                OnDocumentOwnerChanged(previousValue, document);
            }
        }

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected virtual void OnDocumentOwnerChanged(DocX previousValue, DocX newValue)
        {
        }

        /// <summary>
        /// Default constructor
        /// </summary>
        internal DocXElement()
        {
        }

        /// <summary>
        /// Store both the document and xml so that they can be accessed by derived types.
        /// </summary>
        /// <param name="document">The document that this element belongs to.</param>
        /// <param name="xml">The Xml that gives this element substance</param>
        internal DocXElement(DocX document, XElement xml)
        {
            Document = document;
            Xml = xml;
        }
    }
}