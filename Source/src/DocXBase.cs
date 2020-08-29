using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// All DocX types are derived from DocXElement.
    /// This class contains properties which every element of a DocX must contain.
    /// </summary>
    public abstract class DocXBase : IDocumentBase
    {
        private XElement xml;
        private DocX document;

        /// <summary>
        /// The document owner
        /// </summary>
        internal DocX Document
        {
            get => document;
            set
            {
                if (document != value)
                {
                    var previousValue = document;
                    document = value;
                    OnDocumentOwnerChanged(previousValue, document);
                }
            }
        }

        /// <summary>
        /// PackagePart (file) this element is stored in.
        /// </summary>
        internal virtual PackagePart PackagePart { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        internal DocXBase()
        {
        }

        /// <summary>
        /// Store both the document and xml so that they can be accessed by derived types.
        /// </summary>
        /// <param name="document">The document that this element belongs to.</param>
        /// <param name="xml">The Xml that gives this element substance</param>
        internal DocXBase(IDocument document, XElement xml)
        {
            this.xml = xml;
            this.document = (DocX) document;
        }

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
        public IDocument Owner => document;

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected virtual void OnDocumentOwnerChanged(IDocument previousValue, IDocument newValue)
        {
        }
    }
}