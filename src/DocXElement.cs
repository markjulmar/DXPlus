using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// All Document types are derived from DocXElement.
    /// This class contains properties which every element of a Document must contain.
    /// </summary>
    public abstract class DocXElement
    {
        private XElement xml;
        private Document document;
        private PackagePart packagePart;

        /// <summary>
        /// The document owner
        /// </summary>
        internal Document Document
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
        internal PackagePart PackagePart
        {
            get => packagePart;
            set
            {
                if (packagePart != value)
                {
                    var previousValue = packagePart;
                    packagePart = value;
                    OnPackagePartChanged(previousValue, packagePart);
                }
            }
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
        internal DocXElement(IDocument document, XElement xml)
        {
            this.xml = xml;
            this.document = (Document) document;
        }

        /// <summary>
        /// Returns whether this Xml fragment is in a document.
        /// </summary>
        internal bool InDom => Xml?.Parent != null;

        /// <summary>
        /// This is the actual Xml that gives this element substance.
        /// </summary>
        internal XElement Xml
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
        /// Every Document element is connected to a document.
        /// </summary>
        public IDocument Owner => document;

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected virtual void OnDocumentOwnerChanged(IDocument previousValue, IDocument newValue)
        {
        }

        /// <summary>
        /// Called when the package part is changed.
        /// </summary>
        protected virtual void OnPackagePartChanged(PackagePart previousValue, PackagePart newValue)
        {
        }
    }
}