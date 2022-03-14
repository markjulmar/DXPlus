using System.Diagnostics;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Base class for header/footer
/// </summary>
public abstract class HeaderOrFooter : BlockContainer
{
    private XElement? element;

    /// <summary>
    /// This is the actual Xml that gives this element substance.
    /// </summary>
    protected internal override XElement Xml
    {
        get
        {
            if (CreateFunc == null) throw new InvalidOperationException($"{GetType().Name} not created properly.");
            if (!Exists) CreateFunc!.Invoke(this);
            return element!;
        }
        set => element = value;
    }

    /// <summary>
    /// The type of header/footer (even/odd/default)
    /// </summary>
    public HeaderFooterType Type { get; set; }

    /// <summary>
    /// True/False whether the header/footer has been created and exists.
    /// </summary>
    public bool Exists => Id != null && element != null && ExistsFunc?.Invoke(Id)==true;

    /// <summary>
    /// Retrieves the first (main) paragraph for the header/footer.
    /// </summary>
    public Paragraph MainParagraph => Paragraphs.First();

    /// <summary>
    /// Relationship id for the header/footer
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// Get the URI for this header/footer
    /// </summary>
    public Uri Uri => PackagePart.Uri;

    /// <summary>
    /// Constructor
    /// </summary>
    internal HeaderOrFooter()
    {
    }

    // Methods used to create/delete this header/footer.
    internal Action<HeaderOrFooter>? CreateFunc;
    internal Action<HeaderOrFooter>? DeleteFunc;
    internal Func<string, bool>? ExistsFunc;

    /// <summary>
    /// Called to add XML to this element.
    /// </summary>
    /// <param name="xml">XML to add</param>
    /// <returns>Element</returns>
    protected override XElement AddElementToContainer(XElement xml)
    {
        if (CreateFunc == null) throw new InvalidOperationException($"{GetType().Name} not created properly.");

        if (!Exists) CreateFunc.Invoke(this);
        Debug.Assert(element != null);
        return base.AddElementToContainer(xml);
    }

    /// <summary>
    /// Called to add new sections.
    /// </summary>
    /// <param name="breakType"></param>
    /// <returns></returns>
    public override Section AddSection(SectionBreakType breakType) 
        => throw new InvalidOperationException("Cannot add sections to header/footer elements.");

    /// <summary>
    /// Add a new page break to the container
    /// </summary>
    public override void AddPageBreak() 
        => throw new InvalidOperationException("Cannot add page breaks to header/footer elements.");

    /// <summary>
    /// Removes this header/footer
    /// </summary>
    public void Remove()
    {
        if (Exists)
        {
            DeleteFunc!.Invoke(this);
            Id = null;
            element = null;
        }
    }

    /// <summary>
    /// Save the header/footer out to disk.
    /// </summary>
    internal void Save()
    {
        if (Exists)
        {
            PackagePart.Save(new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"), Xml));
        }
    }
}