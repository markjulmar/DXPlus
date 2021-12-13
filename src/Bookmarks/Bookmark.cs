using System;
using DXPlus.Helpers;
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
        /// XML behind this bookmark
        /// </summary>
        internal readonly XElement Xml;

        /// <summary>
        /// Name of the bookmark
        /// </summary>
        public string Name
        {
            get => Xml.AttributeValue(DXPlus.Name.NameId);
            set => Xml.SetAttributeValue(DXPlus.Name.NameId, value);
        }

        /// <summary>
        /// Id for this bookmark in the parent document
        /// </summary>
        public long Id
        {
            get => long.Parse(Xml.AttributeValue(DXPlus.Name.Id));
            set => Xml.SetAttributeValue(DXPlus.Name.Id, value);
        }

        /// <summary>
        /// FirstParagraph this bookmark is part of
        /// </summary>
        public Paragraph Paragraph { get; }

        /// <summary>
        /// Change the text associated with this bookmark.
        /// </summary>
        /// <param name="text">New text value</param>
        public void SetText(string text)
        {
            var nextNode = Xml.NextNode;
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
                AddBookmarkRef(Xml, text);
                return;
            }

            var contentElement = nextElement.Elements(DXPlus.Name.Text).FirstOrDefault();
            if (contentElement == null)
            {
                AddBookmarkRef(Xml, text);
                return;
            }

            contentElement.Value = text;
        }

        /// <summary>
        /// Bookmark constructor
        /// </summary>
        /// <param name="xml">XML for this bookmark</param>
        /// <param name="owner">Associated paragraph object</param>
        public Bookmark(XElement xml, Paragraph owner)
        {
            Xml = xml;
            Paragraph = owner;
        }

        /// <summary>
        /// Adds a bookmark reference
        /// </summary>
        /// <param name="bookmark">Bookmark XML element</param>
        /// <param name="text">Text to insert</param>
        private static void AddBookmarkRef(XNode bookmark, string text)
        {
            if (bookmark == null) 
                throw new ArgumentNullException(nameof(bookmark));
            
            bookmark.AddAfterSelf(HelperFunctions.FormatInput(text, null));
        }
    }
}
