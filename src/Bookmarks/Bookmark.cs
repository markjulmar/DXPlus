using DXPlus.Helpers;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// This represents a bookmark in the Word document.
/// </summary>
public class Bookmark
{
    /// <summary>
    /// XML behind this bookmark
    /// </summary>
    internal readonly XElement Xml;

    /// <summary>
    /// Name of the bookmark
    /// </summary>
    public string Name
    {
        get => Xml.AttributeValue(DXPlus.Name.NameId) ?? throw new DocumentFormatException(nameof(Name));
        set => Xml.SetAttributeValue(DXPlus.Name.NameId, value);
    }

    /// <summary>
    /// Word inserts hidden bookmarks which are prefaced with an underscore. These should be ignored by most processors.
    /// </summary>
    public bool IsHidden => Name.StartsWith('_');

    /// <summary>
    /// Unique identifier for this bookmark in the document
    /// </summary>
    public long Id
    {
        get => long.TryParse(Xml.AttributeValue(DXPlus.Name.Id), out var value) ? value : 0;
        set => Xml.SetAttributeValue(DXPlus.Name.Id, value);
    }

    /// <summary>
    /// Specifies the zero-based index of the first column in this row which shall be part of this bookmark.
    /// This is only used if the bookmark starts and ends within a table.
    /// </summary>
    public int? FirstTableColumn
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "colFirst"), out var val) ? val : null;
        set => Xml.SetAttributeValue(Namespace.Main + "colFirst", value);
    }

    /// <summary>
    /// Specifies the zero-based index of the last column in a row row which shall be part of this bookmark.
    /// This is only used if the bookmark starts and ends within a table.
    /// </summary>
    public int? LastTableColumn
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "colLast"), out var val) ? val : null;
        set => Xml.SetAttributeValue(Namespace.Main + "colLast", value);
    }

    /// <summary>
    /// Paragraph this bookmark is part of
    /// </summary>
    public Paragraph Paragraph { get; }

    /// <summary>
    /// Text associated with this bookmark. Note that this could span across different Run elements.
    /// </summary>
    public string? Text
    {
        get
        {
            var xe = Xml.NextSibling(DXPlus.Name.Run);
            return xe != null ? HelperFunctions.GetText(xe) : null;
        }
    }

    /// <summary>
    /// Change the text associated with this bookmark.
    /// </summary>
    /// <param name="text">New text value</param>
    public bool SetText(string text)
    {
        // Look for either a sibling run, or the bookmarkEnd tag.
        var nextNode = Xml.NextNode;
        if (nextNode == null)
            return false;
        
        var nextElement = nextNode as XElement;
        while (nextElement == null
               || (nextElement.Name != DXPlus.Name.Run && nextElement.Name != DXPlus.Name.BookmarkEnd))
        {
            nextNode = nextNode.NextNode;
            if (nextNode == null)
                return false;

            nextElement = nextNode as XElement;
        }

        // Check if next element is a bookmarkEnd
        if (nextElement.Name == DXPlus.Name.BookmarkEnd)
            return InsertBookmarkText(Xml, text);

        // Or a run
        var contentElement = nextElement.Elements(DXPlus.Name.Text).FirstOrDefault();
        if (contentElement == null)
            return InsertBookmarkText(Xml, text);

        // Set it onto the located text element.
        contentElement.Value = text;
        return true;
    }

    /// <summary>
    /// Bookmark constructor
    /// </summary>
    /// <param name="xml">XML for this bookmark (bookmarkStart)</param>
    /// <param name="owner">Associated paragraph object</param>
    public Bookmark(XElement xml, Paragraph owner)
    {
        Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        Paragraph = owner ?? throw new ArgumentNullException(nameof(owner));
        if (Xml.Name != DXPlus.Name.BookmarkStart)
            throw new ArgumentException($"Cannot create bookmark from {Xml.Name.LocalName}");
    }

    /// <summary>
    /// Adds a bookmark reference
    /// </summary>
    /// <param name="bookmark">Bookmark XML element</param>
    /// <param name="text">Text to insert</param>
    private static bool InsertBookmarkText(XNode bookmark, string text)
    {
        if (bookmark == null) 
            throw new ArgumentNullException(nameof(bookmark));
            
        bookmark.AddAfterSelf(HelperFunctions.FormatInput(text, null));
        
        return true;
    }
}