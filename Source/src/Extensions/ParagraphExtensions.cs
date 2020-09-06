using System;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace DXPlus
{
    /// <summary>
    /// Extension methods to work with the Paragraph type.
    /// </summary>
    public static class ParagraphExtensions
    {
        /// <summary>
        /// Append a new line to this Paragraph.
        /// </summary>
        /// <returns>This Paragraph with a new line appended.</returns>
        public static Paragraph AppendLine(this Paragraph paragraph)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            return paragraph.Append("\n");
        }

        /// <summary>
        /// Append text on a new line to this Paragraph.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="text">The text to append.</param>
        /// <returns>This Paragraph with the new text appended.</returns>
        public static Paragraph AppendLine(this Paragraph paragraph, string text)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            if (text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            return paragraph.Append(text + "\n");
        }

        /// <summary>
        /// Make this paragraph a Heading type by setting the style name
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="headingType"></param>
        public static Paragraph Heading(this Paragraph paragraph, HeadingType headingType)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            paragraph.Properties.StyleName = headingType.GetEnumName();
            return paragraph;
        }

        /// <summary>
        /// Removes characters from a paragraph at a starting index. All
        /// text following that index is removed from the paragraph.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="index">The position to begin deleting characters.</param>
        public static void RemoveText(this Paragraph paragraph, int index)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            paragraph.RemoveText(index, paragraph.Text.Length - index);
        }

        /// <summary>
        /// Fluent method to set StyleName property
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="styleName">Stylename</param>
        public static Paragraph Style(this Paragraph paragraph, string styleName)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            paragraph.Properties.StyleName = styleName;
            return paragraph;
        }

        /// <summary>
        /// Validate that a bookmark exists
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <returns></returns>
        public static bool BookmarkExists(this Paragraph paragraph, string bookmarkName)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            return paragraph.GetBookmarks().Any(b => b.Name.Equals(bookmarkName));
        }

        /// <summary>
        /// Add a paragraph after the current element using the passed text
        /// </summary>
        /// <param name="container">Container owner</param>
        /// <param name="text">Text for new paragraph</param>
        /// <returns>Newly created paragraph</returns>
        public static Paragraph AddParagraph(this InsertBeforeOrAfter container, string text)
        {
            if (container == null)
            {
                throw new ArgumentNullException(nameof(container));
            }

            return container.AddParagraph(text, null);
        }

        /// <summary>
        /// Insert a paragraph before this container.
        /// </summary>
        /// <param name="container">Container owner</param>
        /// <param name="text">Text for new paragraph</param>
        /// <returns>Newly created paragraph</returns>
        public static Paragraph InsertParagraphBefore(this InsertBeforeOrAfter container, string text)
        {
            if (container == null)
            {
                throw new ArgumentNullException(nameof(container));
            }

            return container.InsertParagraphBefore(text, null);
        }
    }
}
