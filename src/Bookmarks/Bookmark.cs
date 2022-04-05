using System.Text;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// This represents a bookmark in the Word document.
/// </summary>
public sealed class Bookmark : XElementWrapper, IEquatable<Bookmark>
{
    /// <summary>
    /// Name of the bookmark
    /// </summary>
    public string Name
    {
        get => Xml.AttributeValue(Internal.Name.NameId) ?? throw new DocumentFormatException(nameof(Name));
        set => Xml!.SetAttributeValue(Internal.Name.NameId, value);
    }

    /// <summary>
    /// Word inserts hidden bookmarks which are prefaced with an underscore. These should be ignored by most processors.
    /// </summary>
    public bool Hidden => Name.StartsWith('_');

    /// <summary>
    /// Unique identifier for this bookmark in the document
    /// </summary>
    public long Id
    {
        get => long.TryParse(Xml.AttributeValue(Internal.Name.Id), out var value) ? value : 0;
        set => Xml!.SetAttributeValue(Internal.Name.Id, value);
    }

    /// <summary>
    /// Specifies the zero-based index of the first column in this row which shall be part of this bookmark.
    /// This is only used if the bookmark starts and ends within a table.
    /// </summary>
    public int? FirstTableColumn
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "colFirst"), out var val) ? val : null;
        set => Xml!.SetAttributeValue(Namespace.Main + "colFirst", value);
    }

    /// <summary>
    /// Specifies the zero-based index of the last column in a row row which shall be part of this bookmark.
    /// This is only used if the bookmark starts and ends within a table.
    /// </summary>
    public int? LastTableColumn
    {
        get => int.TryParse(Xml.AttributeValue(Namespace.Main + "colLast"), out var val) ? val : null;
        set => Xml!.SetAttributeValue(Namespace.Main + "colLast", value);
    }

    /// <summary>
    /// Starting paragraph this bookmark is part of
    /// </summary>
    public Paragraph Paragraph { get; }

    /// <summary>
    /// Text associated with this bookmark. Note that this could span across different Run elements.
    /// </summary>
    public string Text
    {
        get
        {
            var sb = new StringBuilder();
            foreach (var item in Xml!.ElementsAfterSelf())
            {
                if (item.Name.LocalName == Internal.Name.Run.LocalName)
                {
                    sb.Append(DocumentHelpers.GetText(item, false));
                }
                else if (item.Name == Internal.Name.BookmarkEnd
                         && item.AttributeValue(Internal.Name.Id) == Id.ToString())
                    break;
            }

            return sb.ToString();
        }
    }

    /// <summary>
    /// Change the text associated with this bookmark.
    /// </summary>
    /// <param name="text">New text value</param>
    public bool SetText(string text)
    {
        // Look for either a sibling run, or the bookmarkEnd tag.
        var nextNode = Xml!.NextNode;
        if (nextNode == null)
            return false;
        
        var nextElement = nextNode as XElement;
        while (nextElement == null
               || (nextElement.Name != Internal.Name.Run && nextElement.Name != Internal.Name.BookmarkEnd))
        {
            nextNode = nextNode.NextNode;
            if (nextNode == null)
                return false;

            nextElement = nextNode as XElement;
        }

        // Check if next element is a bookmarkEnd
        if (nextElement.Name == Internal.Name.BookmarkEnd)
            return InsertBookmarkText(Xml, text);

        // Or a run
        var contentElement = nextElement.Elements(Internal.Name.Text).FirstOrDefault();
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
        if (Xml.Name != Internal.Name.BookmarkStart)
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
            
        bookmark.AddAfterSelf(DocumentHelpers.CreateRunElements(text, null));
        
        return true;
    }

    /// <summary>
    /// Equality check
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Bookmark? other) 
        => other is not null && (ReferenceEquals(this, other) || Xml!.Equals(other.Xml));

    /// <summary>
    /// Equality check
    /// </summary>
    /// <param name="obj"></param>
    /// <returns></returns>
    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is Bookmark other && Equals(other);

    /// <summary>
    /// Hashcode generator
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml!.GetHashCode();
}