using System;
using System.Collections.Generic;
using System.Linq;

namespace DXPlus
{
    public class BookmarkCollection : List<Bookmark>
    {
        public BookmarkCollection()
        {
        }

        public BookmarkCollection(IEnumerable<Bookmark> bookmarks)
        {
            this.AddRange(bookmarks);
        }

        public Bookmark this[string name] => 
            this.FirstOrDefault(bookmark => string.Equals(bookmark.Name, name, StringComparison.CurrentCultureIgnoreCase));
    }
}
