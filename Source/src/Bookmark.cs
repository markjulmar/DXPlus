using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// This represents a bookmark in the Word document.
    /// </summary>
    public class Bookmark
    {
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Paragraph this bookmark is tied to
        /// </summary>
        public Paragraph Paragraph { get; }

        /// <summary>
        /// Change the text associated with this bookmark.
        /// </summary>
        /// <param name="text">New text value</param>
        public void SetText(string text)
        {
            var bookmark = Paragraph.Xml.Descendants(DXPlus.Name.BookmarkStart)
                                    .FindByAttrVal(DXPlus.Name.NameId, Name);
            if (bookmark == null)
                return;

            var nextNode = bookmark.NextNode;
            var nextElement = nextNode as XElement;
            while (nextElement == null
                   || (nextElement.Name != DXPlus.Name.Run 
                   && nextElement.Name != DXPlus.Name.BookmarkEnd))
            {
                nextNode = nextNode.NextNode;
                nextElement = nextNode as XElement;
            }

            // Check if next element is a bookmarkEnd
            if (nextElement.Name == DXPlus.Name.BookmarkEnd)
            {
                AddBookmarkRef(bookmark, text);
                return;
            }

            var contentElement = nextElement.Elements(DXPlus.Name.Text).FirstOrDefault();
            if (contentElement == null)
            {
                AddBookmarkRef(bookmark, text);
                return;
            }

            contentElement.Value = text;
        }

        /// <summary>
        /// Bookmark constructor
        /// </summary>
        /// <param name="name">Bookmark name</param>
        /// <param name="p">Associated paragraph object</param>
        public Bookmark(string name, Paragraph p)
        {
            Name = name;
            Paragraph = p;
        }

        /// <summary>
        /// Adds a bookmark reference
        /// </summary>
        /// <param name="bookmark">Bookmark XML element</param>
        /// <param name="text">Text to insert</param>
        private void AddBookmarkRef(XNode bookmark, string text)
        {
            var run = HelperFunctions.FormatInput(text, null);
            bookmark.AddAfterSelf(run);
            Paragraph.Document.RenumberIds();
        }
    }
}
