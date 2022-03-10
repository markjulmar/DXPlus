using System.Collections;

namespace DXPlus;

/// <summary>
/// This represents a collection of bookmarks
/// </summary>
public class BookmarkCollection : IReadOnlyList<Bookmark>
{
    private readonly List<Bookmark> bookmarks = new();

    /// <summary>
    /// Bookmark collection initialized with a set of bookmarks
    /// </summary>
    /// <param name="bookmarks">Bookmarks to add to the collection</param>
    public BookmarkCollection(IEnumerable<Bookmark> bookmarks)
    {
        this.bookmarks.AddRange(bookmarks);
    }

    /// <summary>
    /// Index operator to retrieve a bookmark by name
    /// </summary>
    /// <param name="name">Bookmark name</param>
    /// <returns></returns>
    public Bookmark this[string name] =>
        bookmarks.Single(bookmark => string.Equals(bookmark.Name, name, StringComparison.CurrentCultureIgnoreCase));

    /// <summary>
    /// Method to retrieve a bookmark by name.
    /// </summary>
    /// <param name="name">Name to look for</param>
    /// <param name="bookmark">Returning bookmark</param>
    /// <returns>True if found, False if not.</returns>
    public bool TryGetBookmark(string name, out Bookmark? bookmark)
    {
        bookmark = bookmarks.SingleOrDefault(bm => string.Equals(bm.Name, name, StringComparison.CurrentCultureIgnoreCase));
        return bookmark != null;
    }

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<Bookmark> GetEnumerator() => bookmarks.GetEnumerator();

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => ((IEnumerable) bookmarks).GetEnumerator();

    /// <summary>
    /// Gets the number of bookmarks in the collection.
    /// </summary>
    /// <returns>The number of bookmarks in the collection.</returns>
    public int Count => bookmarks.Count;

    /// <summary>
    /// Gets the bookmark at the specified index in the read-only list.
    /// </summary>
    /// <param name="index">The zero-based index of the element to get.</param>
    /// <returns>The bookmark at the specified index in the read-only list.</returns>
    public Bookmark this[int index] => bookmarks[index];
}