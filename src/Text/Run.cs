using System.Diagnostics;
using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a single run of text with optional formatting in a paragraph
/// </summary>
[DebuggerDisplay("{" + nameof(Text) + "}")]
public sealed class Run : DocXElement, IEquatable<Run>
{
    /// <summary>
    /// Retrieves the parent (if any) of this run. This will be a paragraph, hyperlink, etc.
    /// </summary>
    public DocXElement? Parent
    {
        get
        {
            if (!InDocument) return null;

            var parentXml = Xml.Parent;
            while (parentXml != null)
            {
                var wrapper = WrapParentElement(Document, Document.PackagePart, parentXml);
                if (wrapper != null) return wrapper;
                parentXml = parentXml.Parent;
            }

            return null;
        }
    }

    /// <summary>
    /// True if this run has a text block. False if it's a linebreak, paragraph break, or empty.
    /// </summary>
    public bool HasText => Xml.Element(Name.Text) != null;

    /// <summary>
    /// The formatted text value of this run
    /// </summary>
    public string Text => DocumentHelpers.GetText(Xml, false);

    /// <summary>
    /// Returns the breaks in this run
    /// </summary>
    public IEnumerable<ITextElement> Elements
        => Xml.Elements()
            .Where(e => e.Name != Name.RunProperties)
            .Select(WrapTextChild);

    /// <summary>
    /// Wraps a child element in an accessor object.
    /// </summary>
    /// <param name="child"></param>
    /// <returns></returns>
    private ITextElement WrapTextChild(XElement child)
    {
        return child.Name.LocalName switch
        {
            RunTextType.LineBreak => new Break(this, child),
            RunTextType.Text => new Text(this, child),
            RunTextType.DeletedText => new DeletedText(this, child),
            RunTextType.Drawing => new Drawing(Document, Document.PackagePart, child),
            RunTextType.CommentReference => new CommentRef(this, child),
            // Tab, delText, etc.
            _ => new TextElement(this, child),
        };
    }

    /// <summary>
    /// Style applied to this run
    /// </summary>
    public string? StyleName
    {
        get => Xml.GetRunProperties()?.Element(Namespace.Main + "rStyle")?.GetVal();
        set => Xml.CreateRunProperties().AddElementVal(Namespace.Main + "rStyle", value);
    }

    /// <summary>
    /// The run properties for this text run
    /// </summary>
    public Formatting? Properties
    {
        get
        {
            var rPr = Xml.GetRunProperties();
            return rPr == null ? null : new(rPr);
        }
        set
        {
            Xml.GetRunProperties()?.Remove();
            if (value != null)
            {
                var xml = value.Xml!;
                if (xml.Parent != null)
                    xml = xml.Clone();
                Xml.AddFirst(xml);
            }
        }
    }

    /// <summary>
    /// Add/Remove the specific formatting specified from this run.
    /// </summary>
    /// <param name="other">Formatting to apply</param>
    public Run MergeFormatting(Formatting other)
    {
        if (Properties == null) Properties = other;
        else Properties.Merge(other);
        
        return this;
    }

    /// <summary>
    /// Create a run from a string.
    /// </summary>
    /// <param name="text"></param>
    public static implicit operator Run(string text) => new(text);

    /// <summary>
    /// Creates a set of runs from a text string. This properly handles tabs, line breaks, etc.
    /// </summary>
    /// <param name="text">Text</param>
    /// <param name="formatting">Optional formatting</param>
    /// <returns>Enumerable set of run objects</returns>
    public static IEnumerable<Run> Create(string text, Formatting? formatting = null)
    {
        return DocumentHelpers.CreateRunElements(text, formatting?.Xml)
            .Select(xe => new Run(null, null, xe));
    }

    /// <summary>
    /// Public constructor for a run of text
    /// </summary>
    /// <param name="text">Text for this run</param>
    public Run(string text)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));

        var xe = DocumentHelpers.CreateRunElements(text, null).ToList();
        if (xe.Count > 1)
            throw new ArgumentOutOfRangeException(nameof(text),
                "Text cannot mix-in tabs, newlines, or other special characters. Use Run.Create to generate a list of Run objects.");

        Xml = xe.SingleOrDefault() ?? new XElement(Name.Run);
    }

    /// <summary>
    /// Public constructor for a run of text
    /// </summary>
    /// <param name="text">Text for this run</param>
    /// <param name="formatting">Formatting to apply</param>
    public Run(string text, Formatting formatting) : this(text)
    {
        Properties = formatting ?? throw new ArgumentNullException(nameof(formatting));
    }

    /// <summary>
    /// Constructor for a run of text when contained in a document.
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package part this run is in</param>
    /// <param name="xml">XML for the run</param>
    internal Run(Document? document, PackagePart? packagePart, XElement xml) : base(xml)
    {
        if (document != null)
            SetOwner(document, packagePart, false);
    }

    /// <summary>
    /// Split a run at a given index.
    /// </summary>
    /// <param name="index">Index to split this run at</param>
    /// <returns></returns>
    internal (XElement? leftElement, XElement? rightElement) Split(int index)
    {
        // Go through the child (w:t) elements and find the one where this index falls.
        int count = 0, startIndex = 0;
        XElement? textXml = null;
        foreach (var el in Xml.Descendants())
        {
            int size = DocumentHelpers.GetSize(el);
            count += size;
            if (count >= index)
            {
                textXml = el;
                startIndex = count - size;
                break;
            }
        }

        if (textXml == null) return (null, null);

        // Split the block.
        // Returns [textElement, leftSide, rightSide]
        var (leftElement, rightElement) = Split(index, textXml, startIndex);
            
        var splitLeft = new XElement(Xml.Name,
            Xml.Attributes(),
            Xml.Element(Name.RunProperties),
            textXml.ElementsBeforeSelf().Where(n => n.Name != Name.RunProperties),
            leftElement);

        if (DocumentHelpers.GetTextLength(splitLeft) == 0)
        {
            splitLeft = null;
        }

        var splitRight = new XElement(Xml.Name,
            Xml.Attributes(),
            Xml.Element(Name.RunProperties),
            rightElement,
            textXml.ElementsAfterSelf().Where(n => n.Name != Name.RunProperties));

        if (DocumentHelpers.GetTextLength(splitRight) == 0)
        {
            splitRight = null;
        }

        return (splitLeft, splitRight);
    }

    /// <summary>
    /// Split the text block at the given index
    /// </summary>
    /// <param name="index">Index to split at in parent Run</param>
    /// <param name="xml">Text block to split</param>
    /// <param name="startIndex">Start index of the text block in parent Run</param>
    /// <returns>Array with left/right XElement values</returns>
    private static (XElement? leftElement, XElement? rightElement) Split(int index, XElement xml, int startIndex)
    {
        int endIndex = startIndex + DocumentHelpers.GetSize(xml);
        if (index < startIndex || index > endIndex)
            throw new ArgumentOutOfRangeException(nameof(index));

        XElement? splitLeft = null, splitRight = null;

        if (xml.Name.LocalName is RunTextType.Text or RunTextType.DeletedText)
        {
            // The original text element, now containing only the text before the index point.
            splitLeft = new XElement(xml.Name, xml.Attributes(), xml.Value[..(index - startIndex)]);
            if (splitLeft.Value.Length == 0)
            {
                splitLeft = null;
            }
            else
            {
                splitLeft.PreserveSpace();
            }

            // The original text element, now containing only the text after the index point.
            splitRight = new XElement(xml.Name, xml.Attributes(), xml.Value[(index - startIndex)..]);
            if (splitRight.Value.Length == 0)
            {
                splitRight = null;
            }
            else
            {
                splitRight.PreserveSpace();
            }
        }
        else
        {
            if (index == endIndex)
            {
                splitLeft = xml;
            }
            else
            {
                splitRight = xml;
            }
        }

        return (splitLeft, splitRight);
    }

    /// <summary>
    /// Wrap an element
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package part owner</param>
    /// <param name="xml">XML fragment</param>
    /// <returns>Element wrapper</returns>
    private static DocXElement? WrapParentElement(Document document, PackagePart packagePart, XElement xml)
    {
        if (xml.Name.LocalName == Name.Hyperlink.LocalName)
            return new Hyperlink(document, packagePart, xml);
        Debug.Assert(xml.Name == Name.Paragraph);
        var paragraph = document.Paragraphs.FirstOrDefault(p => p.Xml == xml);

        return paragraph;
    }

    /// <summary>
    /// Determines equality for a run
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Run? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a run of text
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Run);

    /// <summary>
    /// Returns hashcode for this run of text
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();

}