using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// Represents a single element contained within the document structure.
/// Wraps the underlying XML element, document owner, and package part.
/// </summary>
public class DocXElement
{
    private XElement? element;
    private Document? document;
    private PackagePart? packagePart;

    /// <summary>
    /// Returns whether this element is in the document structure.
    /// Newly created elements aren't yet in the document and can only be added once.
    /// </summary>
    internal bool InDocument => element?.Parent != null && document != null && packagePart != null;

    /// <summary>
    /// The document owner
    /// </summary>
    internal Document Document => document ?? throw new InvalidOperationException("Element not in a document.");

    /// <summary>
    /// Document owner that doesn't throw
    /// </summary>
    internal Document? SafeDocument => document;

    /// <summary>
    /// Package part that doesn't throw
    /// </summary>
    internal PackagePart? SafePackagePart => packagePart;

    /// <summary>
    /// PackagePart (file) this element is stored in.
    /// </summary>
    internal PackagePart PackagePart => packagePart ?? document?.PackagePart ?? throw new InvalidOperationException("Missing package.");

    /// <summary>
    /// This is the actual Xml that gives this element substance.
    /// </summary>
    public virtual XElement Xml
    {
        get => element ?? throw new InvalidOperationException("Missing XML node.");
        protected internal set => element = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>
    /// This is used to assign an owner to the element. The owner
    /// consists of the owning document object + a zip package.
    /// </summary>
    /// <param name="document">Document</param>
    /// <param name="packagePart">Package</param>
    /// <param name="notify">True to notify about ownership change</param>
    internal void SetOwner(Document document, PackagePart? packagePart, bool notify)
    {
        this.document = document;
        this.packagePart = packagePart;
        if (notify)
        {
            OnAddToDocument();
        }
    }

    /// <summary>
    /// Constructor
    /// </summary>
    internal DocXElement()
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="xml">The XML element that gives this document element substance</param>
    internal DocXElement(XElement xml)
    {
        this.element = xml;
    }

    /// <summary>
    /// This is a reference to the document object that this element belongs to.
    /// It can be null if the given element hasn't been added to a document yet.
    /// </summary>
    public IDocument Owner => Document;

    /// <summary>
    /// Called when the document owner is changed.
    /// </summary>
    protected virtual void OnAddToDocument()
    {
    }
}
