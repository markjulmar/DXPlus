using System.ComponentModel;
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
                var wrapper = WrapDocumentElement(Document, Document.PackagePart, parentXml);
                if (wrapper != null) return wrapper;
                parentXml = parentXml.Parent;
            }

            return null;
        }
    }

    /// <summary>
    /// Gets the start index of this Text (text length before this text)
    /// TODO: remove
    /// </summary>
    public int StartIndex { get; }

    /// <summary>
    /// Gets the end index of this Text (text length before this text + this texts length)
    /// TODO: remove
    /// </summary>
    public int EndIndex { get; }

    /// <summary>
    /// True if this run has a text block. False if it's a linebreak, paragraph break, or empty.
    /// </summary>
    public bool HasText => Xml.Element(Name.Text) != null;
        
    /// <summary>
    /// The formatted text value of this run
    /// </summary>
    public string Text { get; }

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
            "drawing" => new Drawing(Document, Document.PackagePart, child),
            "commentReference" => new CommentRef(this, child),
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
    /// Public constructor for a run of text
    /// </summary>
    /// <param name="text">Text for this run</param>
    public Run(string text)
    {
        Text = text ?? throw new ArgumentNullException(nameof(text));

        var xe = DocumentHelpers.FormatInput(text, null).ToList();
        if (xe.Count > 1)
            throw new InvalidEnumArgumentException(
                "Text cannot mix-in tabs, newlines, or other special characters. Use Run.Create to generate a list of Run objects.");

        Xml = xe.SingleOrDefault() ?? new XElement(Name.Run);
        EndIndex = Text.Length;
    }

    /// <summary>
    /// Creates a set of runs from a text string. This properly handles tabs, line breaks, etc.
    /// </summary>
    /// <param name="text">Text</param>
    /// <param name="formatting">Optional formatting</param>
    /// <returns>Enumerable set of run objects</returns>
    public static IEnumerable<Run> Create(string text, Formatting? formatting = null)
    {
        return DocumentHelpers.FormatInput(text, formatting?.Xml)
            .Select(xe => new Run(null, null, xe, 0));
    }

    /// <summary>
    /// Public constructor for a run of text
    /// </summary>
    /// <param name="text">Text for this run</param>
    /// <param name="formatting">Formatting to apply</param>
    public Run(string text, Formatting formatting)
    {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Xml = new XElement(Name.Run, new XElement(Name.Text, text).PreserveSpace());
        Properties = formatting ?? throw new ArgumentNullException(nameof(formatting));
        EndIndex = Text.Length;
    }

    /// <summary>
    /// Constructor for a run of text when contained in a document.
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package part this run is in</param>
    /// <param name="xml">XML for the run</param>
    /// <param name="startIndex">Start index</param>
    internal Run(Document? document, PackagePart? packagePart, XElement xml, int startIndex) : base(xml)
    {
        if (document != null)
        {
            SetOwner(document, packagePart, false);
        }

        StartIndex = startIndex;
        Text = string.Empty;

        // Determine the end and get the raw text from the run.
        int currentPos = startIndex;
        foreach (var te in xml.Descendants())
        {
            var text = DocumentHelpers.ToText(te);
            if (!string.IsNullOrEmpty(text))
            {
                Text += text;
                currentPos += text.Length;
            }
        }
        EndIndex = currentPos;
    }

    /// <summary>
    /// Split a run at a given index.
    /// </summary>
    /// <param name="index">Index to split this run at</param>
    /// <returns></returns>
    internal XElement?[] SplitAtIndex(int index)
    {
        // Find the (w:t) we need to split based on the index.
        index -= StartIndex;
        var (textXml, startIndex) = FindTextElementByIndex(Xml, index);

        // Split the block.
        // Returns [textElement, leftSide, rightSide]
        var splitText = SplitTextElementAtIndex(index, textXml, startIndex);
            
        var splitLeft = new XElement(Xml.Name,
            Xml.Attributes(),
            Xml.Element(Name.RunProperties),
            splitText[0]!.ElementsBeforeSelf().Where(n => n.Name != Name.RunProperties),
            splitText[1]);

        if (DocumentHelpers.GetTextLength(splitLeft) == 0)
        {
            splitLeft = null;
        }

        var splitRight = new XElement(Xml.Name,
            Xml.Attributes(),
            Xml.Element(Name.RunProperties),
            splitText[2],
            splitText[0]!.ElementsAfterSelf().Where(n => n.Name != Name.RunProperties));

        if (DocumentHelpers.GetTextLength(splitRight) == 0)
        {
            splitRight = null;
        }

        return new[] { splitLeft, splitRight };
    }

    /// <summary>
    /// Split the text block at the given index
    /// </summary>
    /// <param name="index">Index to split at in parent Run</param>
    /// <param name="xml">Text block to split</param>
    /// <param name="startIndex">Start index of the text block in parent Run</param>
    /// <returns>Array with left/right XElement values</returns>
    private static XElement?[] SplitTextElementAtIndex(int index, XElement xml, int startIndex)
    {
        if (xml == null)
            throw new ArgumentNullException(nameof(xml));

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
            splitRight = new XElement(xml.Name, xml.Attributes(), xml.Value.Substring(index - startIndex, xml.Value.Length - (index-startIndex)));
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

        return new[] { xml, splitLeft, splitRight };
    }

    /// <summary>
    /// Internal method to recursively walk all (w:t) elements in this run and find the spot
    /// where an edit (insert/delete) would occur.
    /// </summary>
    /// <param name="element">XML graph to examine</param>
    /// <param name="index">Index to search for</param>
    private static (XElement textXml, int startIndex) FindTextElementByIndex(XElement element, int index)
    {
        int count = 0;
        foreach (var child in element.Descendants())
        {
            int size = DocumentHelpers.GetSize(child);
            count += size;
            if (count >= index)
            {
                return (child, count - size);
            }
        }
        return default;
    }

    /// <summary>
    /// Wrap an element
    /// </summary>
    /// <param name="document">Document owner</param>
    /// <param name="packagePart">Package part owner</param>
    /// <param name="xml">XML fragment</param>
    /// <returns>Element wrapper</returns>
    private static DocXElement? WrapDocumentElement(Document document, PackagePart packagePart, XElement xml)
    {
        if (xml.Name.LocalName == Name.Hyperlink.LocalName)
            return new Hyperlink(document, packagePart, xml);

        if (xml.Name == Name.Paragraph)
        {
            // See if we can find it first. That way we get the proper index.
            var p = document.Paragraphs.FirstOrDefault(p => p.Xml == xml);
            if (p != null) return p;

            // Hmm. Unowned perhaps?
            int pos = 0;
            return DocumentHelpers.WrapParagraphElement(xml, document, packagePart, ref pos);
        }
        return null;
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