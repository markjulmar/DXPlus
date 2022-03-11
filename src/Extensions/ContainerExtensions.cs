using System.Text.RegularExpressions;

namespace DXPlus;

/// <summary>
/// Extension methods for the IContainer type.
/// </summary>
public static class ContainerExtensions
{
    /// <summary>
    /// Add an empty paragraph at the end of this container.
    /// </summary>
    /// <returns>Newly added paragraph</returns>
    public static Paragraph AddParagraph(this IContainer container) => container.AddParagraph(string.Empty);

    /// <summary>
    /// Add a new paragraph with the given text at the end of this container.
    /// </summary>
    /// <returns>Newly added paragraph</returns>
    public static Paragraph AddParagraph(this IContainer container, string text) => container.AddParagraph(text, null);

    /// <summary>
    /// Insert a paragraph with the given text at the specified paragraph index.
    /// </summary>
    /// <param name="container">Container to insert into</param>
    /// <param name="index">Index to insert new paragraph at</param>
    /// <param name="text">Text for new paragraph</param>
    /// <returns>Newly added paragraph</returns>
    public static Paragraph InsertParagraph(this IContainer container, int index, string text) => container.InsertParagraph(index, text, null);

    /// <summary>
    /// Add a new equation using the specified text at the end of this container.
    /// </summary>
    /// <param name="container">Container to add equation to</param>
    /// <param name="equation">Equation</param>
    /// <returns>Newly added paragraph</returns>
    public static Paragraph AddEquation(this IContainer container, string equation) => container.AddParagraph().AppendEquation(equation);

    /// <summary>
    /// Add a new table to the end of this container
    /// </summary>
    /// <param name="container">Container owner</param>
    /// <param name="rows">Rows to add</param>
    /// <param name="columns">Columns to add</param>
    /// <returns></returns>
    public static Table AddTable(this IContainer container, int rows, int columns) => container.AddTable(new Table(rows, columns));

    /// <summary>
    /// Find all occurrences of a string in the paragraph
    /// </summary>
    /// <param name="container"></param>
    /// <param name="text"></param>
    /// <param name="comparisonType"></param>
    /// <returns></returns>
    public static IEnumerable<int> FindText(this IContainer container, string text, StringComparison comparisonType)
    {
        return from p in container.Paragraphs
            from index in p.FindAll(text, comparisonType)
            select index + p.StartIndex;
    }

    /// <summary>
    /// Find all unique instances of the given Regex Pattern,
    /// returning the list of the unique strings found
    /// </summary>
    /// <param name="container"></param>
    /// <param name="regex">Pattern to search for</param>
    /// <returns>Index and matched strings</returns>
    public static IEnumerable<(int index, string text)> FindText(this IContainer container, Regex regex)
    {
        foreach (var p in container.Paragraphs)
        {
            foreach (var (index, text) in p.FindPattern(regex))
            {
                yield return (index: index + p.StartIndex, text);
            }
        }
    }

    /// <summary>
    /// Helper to create a video + thumbnail to insert into the document.
    /// </summary>
    /// <param name="document">Document</param>
    /// <param name="imageFile">Thumbnail</param>
    /// <param name="video">Video URL</param>
    /// <param name="width">Width</param>
    /// <param name="height">Height</param>
    /// <returns>Drawing object to insert</returns>
    public static Drawing? CreateVideo(this IDocument document, string imageFile, Uri video, double width, double height)
    {
        var img = document.AddImage(imageFile);
        var drawing = img.CreatePicture(width,height);
        if (drawing == null) return null;
            
        drawing.Hyperlink = video;
            
        var pic = drawing.Picture;
        if (pic != null)
        {
            pic.Hyperlink = video;
            pic.ImageExtensions.Add(new VideoExtension(video.OriginalString, width, height));
        }

        return drawing;
    }

    /// <summary>
    /// Helper to create a video + thumbnail to insert into the document.
    /// </summary>
    /// <param name="document">Document</param>
    /// <param name="image">Thumbnail</param>
    /// <param name="imageContentType">Image content type</param>
    /// <param name="video">Video URL</param>
    /// <param name="width">Width</param>
    /// <param name="height">Height</param>
    /// <returns>Drawing object to insert</returns>
    public static Drawing? CreateVideo(this IDocument document, Stream image, string imageContentType, Uri video,
        int width, int height)
    {
        var img = document.AddImage(image, imageContentType);
        var drawing = img.CreatePicture(width,height);
        if (drawing == null) return null;

        drawing.Hyperlink = video;
            
        var pic = drawing.Picture;
        if (pic != null)
        {
            pic.Hyperlink = video;
            pic.ImageExtensions.Add(new VideoExtension(video.OriginalString, width, height));
        }

        return drawing;
    }
}