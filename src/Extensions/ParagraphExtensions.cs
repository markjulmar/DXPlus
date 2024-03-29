﻿using System.Drawing;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Extension methods to work with the Paragraph type.
/// </summary>
public static class ParagraphExtensions
{
    /// <summary>
    /// The caption sequence
    /// </summary>
    const string FigureSequence = @" SEQ Figure \* ARABIC ";

    /// <summary>
    /// Append a new line to this Paragraph.
    /// </summary>
    /// <returns>This Paragraph with a new line appended.</returns>
    public static Paragraph Newline(this Paragraph paragraph)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        return paragraph.AddText("\n");
    }

    /// <summary>
    /// Removes characters from a paragraph at a starting index. All
    /// text following that index is removed from the paragraph.
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="index">The position to begin deleting characters.</param>
    public static void RemoveText(this Paragraph paragraph, int index)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        paragraph.RemoveText(index, paragraph.Text.Length - index);
    }

    /// <summary>
    /// Fluent method to set the style from a defined list.
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="headingType"></param>
    public static Paragraph Style(this Paragraph paragraph, HeadingType headingType) 
        => Style(paragraph, headingType.GetEnumName());

    /// <summary>
    /// Fluent method to set the style from a style object
    /// </summary>
    public static Paragraph Style(this Paragraph paragraph, Style style) 
        => Style(paragraph, style.Id ?? throw new ArgumentException("Passed style missing identifier."));

    /// <summary>
    /// Fluent method to set style to a custom name.
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="styleId">Style identifier</param>
    public static Paragraph Style(this Paragraph paragraph, string styleId)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        if (string.IsNullOrWhiteSpace(styleId)) throw new ArgumentException("Value cannot be null or whitespace.", nameof(styleId));

        paragraph.Properties.StyleName = styleId;
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
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        return paragraph.GetBookmarks().Any(b => b.Name.Equals(bookmarkName));
    }

    /// <summary>
    /// Add an empty paragraph after the current element.
    /// </summary>
    /// <returns>Created empty paragraph</returns>
    public static Paragraph AddParagraph(this Block container) => container.AddParagraph(string.Empty);

    /// <summary>
    /// Add a paragraph after the current element using the passed text
    /// </summary>
    /// <param name="container">Container owner</param>
    /// <param name="text">Text for new paragraph</param>
    /// <returns>Newly created paragraph</returns>
    public static Paragraph AddParagraph(this Block container, string text)
    {
        if (container == null) throw new ArgumentNullException(nameof(container));
        if (text == null) throw new ArgumentNullException(nameof(text));
        return container.InsertAfter(new Paragraph(text));
    }

    /// <summary>
    /// Returns any caption on this image.
    /// </summary>
    /// <param name="drawing">Drawing</param>
    /// <returns>Related caption</returns>
    public static string? GetCaption(this Drawing drawing)
    {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (drawing.Parent?.Parent is not Paragraph previousParagraph) return null;

        // Should be in following paragraph.
        var p = previousParagraph.NextParagraph;
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
        if (string.IsNullOrEmpty(captionText)) throw new ArgumentException("Value cannot be null or empty.", nameof(captionText));
        if (!drawing.InDocument) throw new InvalidOperationException("Drawing must be part of a document.");
        if (drawing.Parent?.Parent is not Paragraph paragraphOwner) throw new InvalidOperationException("Drawing must be in paragraph to add caption.");

        if (paragraphOwner.NextParagraph?.Xml.Descendants(Name.SimpleField)
                .FirstOrDefault(x => x.Attribute(Name.Instr)?.Value == FigureSequence) != null)
            throw new ArgumentException("Drawing already has caption.", nameof(drawing));

        var document = drawing.Document;
        var captionStyle = document.Styles.Find(HeadingType.Caption.ToString(), StyleType.Paragraph);
        if (captionStyle == null)
        {
            captionStyle = document.Styles.Add(HeadingType.Caption.ToString(), HeadingType.Caption.ToString(), StyleType.Paragraph);
            captionStyle.NextParagraphStyle = "Normal";
            captionStyle.ParagraphFormatting = new()
            {
                LineSpacingAfter = Uom.FromPoints(10).Dxa,
                LineSpacing = Uom.FromPoints(12).Dxa,
                LineRule = LineRule.Auto,
                DefaultFormatting =
                {
                    Italic = true,
                    Color = Color.FromArgb(0x1F,0x37,0x63),
                    FontSize = 9
                }
            };
        }

        var captionParagraph = new Paragraph().Style(HeadingType.Caption.ToString());
        captionParagraph.AddText("Figure ");

        var figureIds =
            document.Xml.Descendants(Name.SimpleField)
                .Where(x => x.Attribute(Name.Instr)?.Value == FigureSequence)
                .Select(x => x.Descendants(Name.Text).SingleOrDefault())
                .Select(x => int.TryParse(x?.Value??"0", out var n) ? n : -1).ToList();

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
            
        captionParagraph.AddText(captionText);

        paragraphOwner.Xml.AddAfterSelf(captionParagraph.Xml);
        return captionParagraph;
    }

    /// <summary>
    /// Set all the outside borders of the table
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="border"></param>
    /// <returns></returns>
    public static Paragraph SetOutsideBorders(this Paragraph paragraph, Border border)
    {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        paragraph.Properties.SetOutsideBorders(border);
        return paragraph;
    }

    /// <summary>
    /// Set all the outside borders of the table
    /// </summary>
    /// <param name="properties"></param>
    /// <param name="border"></param>
    /// <returns></returns>
    public static void SetOutsideBorders(this ParagraphProperties properties, Border border)
    {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        properties.LeftBorder = border;
        properties.RightBorder = border;
        properties.TopBorder = border;
        properties.BottomBorder = border;
    }

    /// <summary>
    /// Attaches a comment to this paragraph.
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="comment">Comment to add</param>
    public static void AttachComment(this Paragraph paragraph, Comment comment) => paragraph.AttachComment(comment, paragraph.Runs.First());

    /// <summary>
    /// Attach a comment to this Run
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="comment">Comment</param>
    /// <param name="run">Text run</param>
    public static void AttachComment(this Paragraph paragraph, Comment comment, Run run) => paragraph.AttachComment(comment, run, run);
}