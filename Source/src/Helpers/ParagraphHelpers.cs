using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Fluent methods for the paragraph type
    /// </summary>
    public static class ParagraphHelpers
    {
        /// <summary>
        /// Fluent method to set alignment
        /// </summary>
        /// <param name="paragraph">Paragraph</param>
        /// <param name="alignment">Desired alignment</param>
        public static Paragraph Align(this Paragraph paragraph, Alignment alignment)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Alignment = alignment;
            return paragraph;
        }

        /// <summary>
        /// Append a new line to this Paragraph.
        /// </summary>
        /// <returns>This Paragraph with a new line appended.</returns>
        public static Paragraph AppendLine(this Paragraph paragraph)
        {
            if (paragraph == null) 
                throw new ArgumentNullException(nameof(paragraph));

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
                throw new ArgumentNullException(nameof(paragraph));
            if (text == null) 
                throw new ArgumentNullException(nameof(text));
            
            return paragraph.Append("\n" + text);
        }

        /// <summary>
        /// Make the prior inserted text bold.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>This Paragraph with the last appended text bold.</returns>
        public static Paragraph Bold(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Bold = true;
            return paragraph;
        }

        /// <summary>
        /// Append text to this Paragraph and then set it to full caps.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="capsStyle">The caps style to apply to the last appended text.</param>
        /// <returns>This Paragraph with the last appended text's caps style changed.</returns>
        public static Paragraph CapsStyle(this Paragraph paragraph, CapsStyle capsStyle)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.CapsStyle = capsStyle;
            return paragraph;
        }

        /// <summary>
        /// Apply a color to the text runs in this paragraph.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="color">A color to use on the appended text.</param>
        /// <returns>This Paragraph with the last appended text colored.</returns>
        public static Paragraph Color(this Paragraph paragraph, Color color)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Color = color;
            return paragraph;
        }

        /// <summary>
        /// Set the culture of the preceding text
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="culture">The CultureInfo for text</param>
        /// <returns>This Paragraph in current culture</returns>
        public static Paragraph Culture(this Paragraph paragraph, CultureInfo culture)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Culture = culture;
            return paragraph;
        }

        /// <summary>
        /// Set the culture of the preceding text
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>This Paragraph in current culture</returns>
        public static Paragraph Culture(this Paragraph paragraph) => Culture(paragraph, CultureInfo.CurrentCulture);

        /// <summary>
        /// Set the font for the preceding text.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="fontFamily">The font to use for the appended text.</param>
        /// <returns>This Paragraph with the last appended text's font changed.</returns>
        public static Paragraph Font(this Paragraph paragraph, FontFamily fontFamily)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Font = fontFamily;
            return paragraph;
        }

        /// <summary>
        /// Set the font size for the appended text in this paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="fontSize">The font size to use for the appended text.</param>
        /// <returns>Paragraph with the last appended text resized.</returns>
        public static Paragraph FontSize(this Paragraph paragraph, double fontSize)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.FontSize = fontSize;
            return paragraph;
        }

        /// <summary>
        /// Make this paragraph a Heading type by setting the style name
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="headingType"></param>
        public static Paragraph Heading(this Paragraph paragraph, HeadingType headingType)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.StyleName = headingType.GetEnumName();
            return paragraph;
        }

        /// <summary>
        /// Append text to this Paragraph and then make it italic.
        /// </summary>
        /// <returns>This Paragraph with the last appended text italic.</returns>
        public static Paragraph Italic(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Italic = true;
            return paragraph;
        }

        /// <summary>
        /// Hide or show this paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="hide">True to hide</param>
        /// <returns>This Paragraph with the last appended text hidden.</returns>
        public static Paragraph Hide(this Paragraph paragraph, bool hide = true)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.IsHidden = hide;
            return paragraph;
        }

        /// <summary>
        /// Highlights the given paragraph/line
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="highlight">The highlight to apply to the last appended text.</param>
        /// <returns>This Paragraph with the last appended text highlighted.</returns>
        public static Paragraph Highlight(this Paragraph paragraph, Highlight highlight)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Highlight = highlight;
            return paragraph;
        }

        /// <summary>
        /// Set the kerning value for the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="kerning"></param>
        /// <returns></returns>
        public static Paragraph Kerning(this Paragraph paragraph, int kerning)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Kerning = kerning;
            return paragraph;
        }

        /// <summary>
        /// Set one of the misc. properties
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="effect">The miscellaneous property to set.</param>
        /// <returns>This Paragraph with the last appended text changed by a miscellaneous property.</returns>
        public static Paragraph Effect(this Paragraph paragraph, Effect effect)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Effect = effect;
            return paragraph;
        }

        /// <summary>
        /// Set the percentage scale for the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="scale">% to expand/compress the run</param>
        /// <returns>Paragraph</returns>
        public static Paragraph ExpansionScale(this Paragraph paragraph, int scale)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.ExpansionScale = scale;
            return paragraph;
        }

        /// <summary>
        /// Set the vertical position of the paragraph in half-points
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="position"></param>
        /// <returns>Paragraph</returns>
        public static Paragraph Position(this Paragraph paragraph, double position)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));
            
            paragraph.Position = position;
            return paragraph;
        }

        /// <summary>
        /// Removes characters from a DXPlus.DocX.Paragraph.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="index">The position to begin deleting characters.</param>
        /// <param name="trackChanges">Track changes</param>
        public static void RemoveText(this Paragraph paragraph, int index, bool trackChanges = false)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.RemoveText(index, paragraph.Text.Length - index, trackChanges);
        }

        /// <summary>
        /// Append text to this Paragraph and then set it to superscript.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>This Paragraph with the last appended text's script style changed.</returns>
        public static Paragraph Superscript(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Superscript = true;
            return paragraph;
        }

        /// <summary>
        /// Append text to this Paragraph and then set it to subscript.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>This Paragraph with the last appended text's script style changed.</returns>
        public static Paragraph Subscript(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.Subscript = true;
            return paragraph;
        }

        /// <summary>
        /// Set the spacing for the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="spacing"></param>
        /// <returns></returns>
        public static Paragraph LineSpacing(this Paragraph paragraph, double spacing)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.LineSpacing = spacing;
            return paragraph;
        }

        /// <summary>
        /// Set the spacing after the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="spacing"></param>
        /// <returns></returns>
        public static Paragraph LineSpacingAfter(this Paragraph paragraph, double spacing)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));
            
            paragraph.LineSpacingAfter = spacing;
            return paragraph;
        }

        /// <summary>
        /// Set the spacing before the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="spacing"></param>
        /// <returns></returns>
        public static Paragraph LineSpacingBefore(this Paragraph paragraph, double spacing)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.LineSpacingBefore = spacing;
            return paragraph;
        }

        /// <summary>
        /// Adds a single or double-line strikethrough on the text.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="strike">The strike through style to used on the last appended text.</param>
        /// <returns>This Paragraph with the last appended text striked.</returns>
        public static Paragraph StrikeThrough(this Paragraph paragraph, Strikethrough strike)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.StrikeThrough = strike;
            return paragraph;
        }

        /// <summary>
        /// Fluent method to set StyleName property
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="styleName">Stylename</param>
        public static Paragraph Style(this Paragraph paragraph, string styleName)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.StyleName = styleName;
            return paragraph;
        }

        /// <summary>
        /// Append text to this Paragraph and then underline it.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="underlineStyle">The underline style to use for the appended text.</param>
        /// <returns>This Paragraph with the last appended text underlined.</returns>
        public static Paragraph UnderlineStyle(this Paragraph paragraph, UnderlineStyle underlineStyle)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            paragraph.UnderlineStyle = underlineStyle;
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
            return paragraph.GetBookmarks().Any(b => b.Name.Equals(bookmarkName));
        }
    }
}
