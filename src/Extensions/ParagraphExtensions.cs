using System;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Extension methods to work with the FirstParagraph type.
    /// </summary>
    public static class ParagraphExtensions
    {
        /// <summary>
        /// The caption sequence
        /// </summary>
        const string FigureSequence = @" SEQ Figure \* ARABIC ";

        /// <summary>
        /// Append a new line to this FirstParagraph.
        /// </summary>
        /// <returns>This FirstParagraph with a new line appended.</returns>
        public static Paragraph AppendLine(this Paragraph paragraph)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            return paragraph.Append("\n");
        }

        /// <summary>
        /// Append text on a new line to this FirstParagraph.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="text">The text to append.</param>
        /// <param name="formatting">Optional formatting for text</param>
        /// <returns>This FirstParagraph with the new text appended.</returns>
        public static Paragraph AppendLine(this Paragraph paragraph, string text, Formatting formatting = null)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            if (text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            return paragraph.Append(text + "\n", formatting);
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
        /// Fluent method to set the style from a defined list.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="headingType"></param>
        public static Paragraph Style(this Paragraph paragraph, HeadingType headingType) => Style(paragraph, headingType.GetEnumName());

        /// <summary>
        /// Fluent method to set the style from a style object
        /// </summary>
        public static Paragraph Style(this Paragraph paragraph, Style style) => Style(paragraph, style.Id);

        /// <summary>
        /// Fluent method to set style to a custom name.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="styleName">Stylename</param>
        public static Paragraph Style(this Paragraph paragraph, string styleName)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));
            if (string.IsNullOrWhiteSpace(styleName))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(styleName));

            paragraph.Properties.StyleName = styleName;
            return paragraph;
        }

        /// <summary>
        /// Sets a specific border on the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="type"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <param name="spacing"></param>
        /// <param name="size"></param>
        /// <param name="shadow"></param>
        /// <returns></returns>
        public static Paragraph WithBorder(this Paragraph paragraph, ParagraphBorderType type, BorderStyle style, Color color, double? spacing = 1,
                                    double size = 2, bool shadow = false)
        {
            paragraph.Properties.SetBorder(type, style, color, spacing, size, shadow);
            return paragraph;
        }

        /// <summary>
        /// Sets all borders on the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="type"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <param name="spacing"></param>
        /// <param name="size"></param>
        /// <param name="shadow"></param>
        /// <returns></returns>
        public static Paragraph WithBorders(this Paragraph paragraph, BorderStyle style, Color color, double? spacing = 1,
            double size = 2, bool shadow = false)
        {
            paragraph.Properties.SetBorders(style, color, spacing, size, shadow);
            return paragraph;
        }

        /// <summary>
        /// Set a fill color on the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="fillColor"></param>
        /// <returns></returns>
        public static Paragraph WithFill(this Paragraph paragraph, Color fillColor)
        {
            paragraph.Properties.ShadeFill = fillColor;
            return paragraph;
        }

        /// <summary>
        /// Set a fill color on the paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="pattern"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static Paragraph WithPattern(this Paragraph paragraph, ShadePattern pattern, Color color)
        {
            paragraph.Properties.ShadePattern = pattern;
            paragraph.Properties.ShadeColor = color;
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
        public static Paragraph AddParagraph(this Block container, string text)
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
        public static Paragraph InsertParagraphBefore(this Block container, string text)
        {
            if (container == null)
            {
                throw new ArgumentNullException(nameof(container));
            }

            return container.InsertParagraphBefore(text, null);
        }

        /// <summary>
        /// Returns any caption on this image.
        /// </summary>
        /// <param name="drawing">Drawing</param>
        /// <returns>Related caption</returns>
        public static string GetCaption(this Drawing drawing)
        {
            if (drawing == null) throw new ArgumentNullException(nameof(drawing));
            if (drawing.Parent?.Parent is not Paragraph previousParagraph) return null;

            // Should be in following paragraph.
            var p = previousParagraph.NextParagagraph;
            return p?.Properties.StyleName == "Caption" ? p.Text : null;
        }

        /// <summary>
        /// Adds a caption in the form of "Figure n {captionText}".
        /// </summary>
        /// <param name="drawing">Drawing to insert the caption for</param>
        /// <param name="captionText">Text to add</param>
        /// <returns>Paragraph with caption</returns>
        public static Paragraph AddCaption(this Drawing drawing, string captionText)
        {
            if (drawing == null) throw new ArgumentNullException(nameof(drawing));
            if (string.IsNullOrEmpty(captionText))
                throw new ArgumentException("Value cannot be null or empty.", nameof(captionText));

            if (!drawing.InDom || drawing.Parent?.Parent is not Paragraph previousParagraph)
                throw new ArgumentException("Drawing must be in document.", nameof(drawing));

            if (previousParagraph.NextParagagraph?.Xml.Descendants(Name.SimpleField)
                    .FirstOrDefault(x => x.Attribute(Name.Instr)?.Value == FigureSequence) != null)
                throw new ArgumentException("Drawing already has caption.", nameof(drawing));

            var document = drawing.Document;
            var captionStyle = document.Styles.GetStyle(HeadingType.Caption.ToString(), StyleType.Paragraph);
            if (captionStyle == null)
            {
                captionStyle = document.Styles.AddStyle(HeadingType.Caption.ToString(), StyleType.Paragraph);
                captionStyle.NextParagraphStyle = "Normal";
                captionStyle.ParagraphFormatting.LineSpacingAfter = Uom.FromPoints(10).Dxa;
                captionStyle.ParagraphFormatting.LineSpacing = Uom.FromPoints(12).Dxa;
                captionStyle.ParagraphFormatting.LineRule = LineRule.Auto;
                captionStyle.ParagraphFormatting.DefaultFormatting.Italic = true;
                captionStyle.ParagraphFormatting.DefaultFormatting.Color = Color.FromArgb(0x1F,0x37,0x63);
                captionStyle.ParagraphFormatting.DefaultFormatting.FontSize = 9;
            }

            var captionParagraph = new Paragraph().Style(HeadingType.Caption.ToString());
            captionParagraph.Append("Figure ");

            var figureIds =
                document.Xml.Descendants(Name.SimpleField)
                    .Where(x => x.Attribute(Name.Instr)?.Value == FigureSequence)
                    .Select(x => x.Descendants(Name.Text).SingleOrDefault())
                    .Select(x => int.TryParse(x?.Value??"0", out var n) == true ? n : -1).ToList();

            int nextNum = figureIds.Count > 0 ? figureIds.Max()+1 : 1;
            if (nextNum <= 0)
                nextNum = 1;

            captionParagraph.Xml.Add(new XElement(Name.SimpleField,
                new XAttribute(Name.Instr, FigureSequence),
                new XElement(Name.Run,
                    new XElement(Name.RunProperties,
                        new XElement(Namespace.Main + "noProof")),
                    new XElement(Name.Text, nextNum))));

            // Add a space if there's no separator.
            if (captionText[0] != ':' && captionText[0] != ' ')
                captionText = " " + captionText;
            
            captionParagraph.Append(captionText);

            previousParagraph.Xml.AddAfterSelf(captionParagraph.Xml);
            return captionParagraph;
        }
    }
}
