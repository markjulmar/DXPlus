namespace DXPlus;

/// <summary>
/// This represents a collection of bookmarks
/// </summary>
public class BookmarkCollection : List<Bookmark>
{
    /// <summary>
    /// Empty bookmark collection
    /// </summary>
    public BookmarkCollection()
    {
    }

    /// <summary>
    /// Bookmark collection initialized with a set of bookmarks
    /// </summary>
    /// <param name="bookmarks">Bookmarks to add to the collection</param>
    public BookmarkCollection(IEnumerable<Bookmark> bookmarks)
    {
        AddRange(bookmarks);
    }

    /// <summary>
    /// Index operator to retrieve a bookmark by name
    /// </summary>
    /// <param name="name">Bookmark name</param>
    /// <returns></returns>
    public Bookmark this[string name] =>
        this.FirstOrDefault(bookmark => 
            string.Equals(bookmark.Name, name, StringComparison.CurrentCultureIgnoreCase));
}