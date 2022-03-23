using System.IO.Packaging;

namespace DXPlus;

/// <summary>
/// Represents an Image embedded in a document.
/// </summary>
public class Image
{
    private readonly Document? document;
    private readonly PackageRelationship? packageRelationship;

    /// <summary>
    /// Get the stream for the picture
    /// </summary>
    /// <returns>Open stream</returns>
    public Stream OpenStream()
    {
        string? targetUrl = Url;
        if (targetUrl == null)
            throw new InvalidOperationException("Image stream not available.");

        // Return a readonly stream to the image
        return document!.Package.GetPart(new Uri(targetUrl, UriKind.Relative))
            .GetStream(FileMode.Open, FileAccess.Read);
    }

    /// <summary>
    /// Uri for this image
    /// </summary>
    public Uri? Uri => packageRelationship?.TargetUri;

    /// <summary>
    /// URL for the image contained within the package.
    /// </summary>
    private string? Url
    {
        get
        {
            string? targetUrl = packageRelationship?.TargetUri.OriginalString;
            if (targetUrl != null && targetUrl[0] != '/')
            {
                string temp = packageRelationship!.SourceUri.OriginalString;
                string start = temp.Remove(temp.LastIndexOf('/'));
                targetUrl = start + "/" + targetUrl;
            }
            return targetUrl;
        }
    }

    /// <summary>
    /// Returns the id of this Image.
    /// </summary>
    public string Id { get; }

    ///<summary>
    /// Returns the name of the image file.
    ///</summary>
    public string FileName => Uri != null ? Path.GetFileName(Uri!.ToString()) : string.Empty;

    /// <summary>
    /// Returns the image type (extension) such as .png, .jpg, .svg, etc.
    /// </summary>
    public string ImageType => Path.GetExtension(FileName).ToLower();

    /// <summary>
    /// Internal constructor to create an image from a package relationship.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="packageRelationship"></param>
    /// <param name="id">Optional package identifier</param>
    internal Image(Document? document, PackageRelationship? packageRelationship, string? id = null)
    {
        this.document = document;
        this.packageRelationship = packageRelationship;
        Id = id ?? packageRelationship?.Id ?? string.Empty;
    }

    /// <summary>
    /// Create a new picture which can be added to a paragraph from this image.
    /// </summary>
    /// <param name="name">Name of the image</param>
    /// <param name="description">Description of the image</param>
    /// <returns>New picture</returns>
    public Picture CreatePicture(string name, string description)
    {
        if (document == null)
            throw new InvalidOperationException("Cannot create picture from image obtained by unowned element.");
        return document.CreateDrawingWithEmbeddedPicture(Id, name, description).Picture!;
    }

    /// <summary>
    /// Create a new picture which can be added to a paragraph from this image.
    /// </summary>
    /// <param name="width">Width of the picture</param>
    /// <param name="height">Height of the picture</param>
    /// <returns>New picture</returns>
    public Picture CreatePicture(double width, double height) =>
        CreatePicture(string.Empty, string.Empty, width, height);

    /// <summary>
    /// Create a new picture with specific dimensions and insert it into a paragraph.
    /// </summary>
    /// <param name="name">Name of the picture</param>
    /// <param name="description">Description of the picture</param>
    /// <param name="height">Rendered height</param>
    /// <param name="width">Rendered width</param>
    /// <returns>New picture</returns>
    public Picture CreatePicture(string name, string description, double width, double height)
    {
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));

        // Create the drawing.
        var picture = CreatePicture(name, description);

        // Set the extent on the drawing
        picture.Width = width;
        picture.Height = height;

        // Use the same values for the drawing owner
        var drawing = picture.Drawing!;
        drawing.Width = width;
        drawing.Height = height;
            
        return picture;
    }
}