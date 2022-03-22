using System.Text.RegularExpressions;
using DXPlus.Internal;

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
    public static Paragraph AddParagraph(this IContainer container) 
        => container.Add(new Paragraph());

    /// <summary>
    /// Add a new equation using the specified text at the end of this container.
    /// </summary>
    /// <param name="container">Container to add equation to</param>
    /// <param name="equation">Equation</param>
    /// <returns>Newly added paragraph</returns>
    public static Paragraph AddEquation(this IContainer container, string equation) 
        => container.AddParagraph().AddEquation(equation);

    /// <summary>
    /// Find all occurrences of a string in the container. This searches headers, all paragraphs, and footers.
    /// </summary>
    /// <param name="container"></param>
    /// <param name="findText"></param>
    /// <param name="comparisonType"></param>
    /// <returns></returns>
    public static IEnumerable<(Paragraph paragraphOwner, int index)> FindText(this IContainer container, string findText, StringComparison comparisonType)
    {
        if (container == null) throw new ArgumentNullException(nameof(container));
        if (string.IsNullOrEmpty(findText)) throw new ArgumentNullException(nameof(findText));

        return container.Sections.SelectMany(s => s.Headers).SelectMany(header => header.Paragraphs)
            .Union(container.Paragraphs)
            .Union(container.Sections.SelectMany(s => s.Footers).SelectMany(footer => footer.Paragraphs))
            .ToList()
            .SelectMany(p => p.FindAll(findText, comparisonType).Select(n => (p, n)));
    }

    /// <summary>
    /// Find all unique instances of the given Regex Pattern,
    /// returning the list of the unique strings found
    /// </summary>
    /// <param name="container"></param>
    /// <param name="regex">Pattern to search for</param>
    /// <returns>Index and matched strings</returns>
    public static IEnumerable<(Paragraph paragraphOwner, int index, string text)> FindText(this IContainer container, Regex regex)
    {
        if (container == null) throw new ArgumentNullException(nameof(container));
        if (regex == null) throw new ArgumentNullException(nameof(regex));

        foreach (var p in container.Sections.SelectMany(s => s.Headers).SelectMany(header => header.Paragraphs)
                     .Union(container.Paragraphs)
                     .Union(container.Sections.SelectMany(s => s.Footers).SelectMany(footer => footer.Paragraphs)))
        {
            foreach (var (index, text) in p.FindPattern(regex))
            {
                yield return (p, index, text);
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
    public static Drawing CreateVideo(this IDocument document, string imageFile, Uri video, double width, double height)
    {
        var img = document.CreateImage(imageFile);
        var drawing = img.CreatePicture(width,height);
            
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
    public static Drawing CreateVideo(this IDocument document, Stream image, string imageContentType, Uri video,
        int width, int height)
    {
        var img = document.CreateImage(image, imageContentType);
        var drawing = img.CreatePicture(width,height);

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