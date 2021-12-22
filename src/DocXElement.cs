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
        protected Document document;
        protected PackagePart packagePart;

        /// <summary>
        /// The document owner
        /// </summary>
        internal virtual Document Document => document;

        /// <summary>
        /// This is used to assign an owner to the element. The owner
        /// consists of the owning document object + a zip package.
        /// </summary>
        /// <param name="document">Document</param>
        /// <param name="packagePart">Package</param>
        internal void SetOwner(Document document, PackagePart packagePart)
        {
            this.document = document;
            this.packagePart = packagePart;

            if (document != null)
            {
                OnDocumentOwnerChanged();
            }
        }

        /// <summary>
        /// PackagePart (file) this element is stored in.
        /// </summary>
        internal virtual PackagePart PackagePart => packagePart;

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
        /// <param name="packagePart">The package this element is in.</param>
        /// <param name="xml">The Xml that gives this element substance</param>
        internal DocXElement(IDocument document, PackagePart packagePart, XElement xml)
        {
            this.packagePart = packagePart;
            this.document = (Document) document;
            Xml = xml;
        }

        /// <summary>
        /// Returns whether this Xml fragment is in a document.
        /// </summary>
        internal bool InDom => Xml?.Parent != null;

        /// <summary>
        /// This is the actual Xml that gives this element substance.
        /// </summary>
        internal virtual XElement Xml { get; set; }

        /// <summary>
        /// This is a reference to the document object that this element belongs to.
        /// It can be null if the given element hasn't been added to a document yet.
        /// </summary>
        public IDocument Owner => Document;

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected virtual void OnDocumentOwnerChanged()
        {
        }
    }
}