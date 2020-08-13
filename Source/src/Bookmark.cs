using System.Linq;
using System.Xml.Linq;

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

        public void SetText(string newText)
        {
            ReplaceBookmark(Name, newText);
        }

        public Bookmark(string name, Paragraph p)
        {
            Name = name;
            Paragraph = p;
        }

        private void AddBookmarkRef(string toInsert, XElement bookmark)
        {
            var run = HelperFunctions.FormatInput(toInsert, null);
            bookmark.AddAfterSelf(run);
            Paragraph.runs = Paragraph.Xml.Elements(DocxNamespace.Main + "r").ToList();
            HelperFunctions.RenumberIDs(Paragraph.Document);
        }

        private void ReplaceBookmark(string bookmarkName, string toInsert)
        {
            XElement bookmark = Paragraph.Xml.Descendants(DocxNamespace.Main + "bookmarkStart")
                                         .FindByAttrVal(DocxNamespace.Main + "name", bookmarkName);
            if (bookmark == null)
                return;

            XNode nextNode = bookmark.NextNode;
            XElement nextElement = nextNode as XElement;
            while (nextElement == null
                    || nextElement.Name.NamespaceName != DocxNamespace.Main.NamespaceName 
                    || (nextElement.Name.LocalName != "r" 
                        && nextElement.Name.LocalName != "bookmarkEnd"))
            {
                nextNode = nextNode.NextNode;
                nextElement = nextNode as XElement;
            }

            // Check if next element is a bookmarkEnd
            if (nextElement.Name.LocalName == "bookmarkEnd")
            {
                AddBookmarkRef(toInsert, bookmark);
                return;
            }

            XElement contentElement = nextElement.Elements(DocxNamespace.Main + "t").FirstOrDefault();
            if (contentElement == null)
            {
                AddBookmarkRef(toInsert, bookmark);
                return;
            }

            contentElement.Value = toInsert;
        }

    }
}
