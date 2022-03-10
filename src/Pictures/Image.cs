using System.IO.Packaging;

namespace DXPlus;

/// <summary>
/// Represents an Image embedded in a document.
/// </summary>
public class Image
{
    private readonly Document? document;

    /// <summary>
    /// Associated package relationship
    /// </summary>
    internal PackageRelationship? PackageRelationship { get; }

    /// <summary>
    /// Get the stream for the picture
    /// </summary>
    /// <returns>Open stream</returns>
    public Stream? OpenStream()
    {
        if (document == null || PackageRelationship == null) 
            return null;

        string targetUrl = PackageRelationship.TargetUri.OriginalString;
        if (targetUrl[0] != '/')
        {
            string temp = PackageRelationship.SourceUri.OriginalString;
            string start = temp.Remove(temp.LastIndexOf('/'));
            targetUrl = start + "/" + targetUrl;
        }

        // Return a readonly stream to the image
        return document.Package.GetPart(new Uri(targetUrl, UriKind.Relative))
            .GetStream(FileMode.Open, FileAccess.Read);
    }

    /// <summary>
    /// Returns the id of this Image.
    /// </summary>
    public string Id { get; }

    ///<summary>
    /// Returns the name of the image file.
    ///</summary>
    public string FileName => Path.GetFileName(PackageRelationship.TargetUri.ToString());

    /// <summary>
    /// Returns the image type (extension) such as .png, .jpg, .svg, etc.
    /// </summary>
    public string ImageType => Path.GetExtension(FileName).ToLower();

    /// <summary>
    /// Internal constructor to create an image from a package relationship.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="packageRelationship"></param>
    internal Image(Document document, PackageRelationship packageRelationship)
    {
        this.document = document ?? throw new ArgumentNullException(nameof(document));
        this.PackageRelationship = packageRelationship ?? throw new ArgumentNullException(nameof(packageRelationship));
        Id = packageRelationship.Id;
    }

    /// <summary>
    /// Internal constructor to create an image from an unowned document
    /// </summary>
    /// <param name="id"></param>
    internal Image(string id)
    {
        this.Id = id;
    }

    /// <summary>
    /// Create a new picture which can be added to a paragraph from this image.
    /// </summary>
    /// <param name="name">Name of the image</param>
    /// <param name="description">Description of the image</param>
    /// <returns>New picture</returns>
    public Drawing CreatePicture(string name, string description) =>
        document.CreateDrawingWithEmbeddedPicture(Id, name, description);

    /// <summary>
    /// Create a new picture which can be added to a paragraph from this image.
    /// </summary>
    /// <param name="width">Width of the picture</param>
    /// <param name="height">Height of the picture</param>
    /// <returns>New picture</returns>
    public Drawing CreatePicture(double width, double height) =>
        CreatePicture(string.Empty, string.Empty, width, height);

    /// <summary>
    /// Create a new picture with specific dimensions and insert it into a paragraph.
    /// </summary>
    /// <param name="name">Name of the picture</param>
    /// <param name="description">Description of the picture</param>
    /// <param name="height">Rendered height</param>
    /// <param name="width">Rendered width</param>
    /// <returns>New picture</returns>
    public Drawing CreatePicture(string name, string description, double width, double height)
    {
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));

        // Create the drawing.
        var drawing = CreatePicture(name, description);
            
        // Set the extent on the drawing
        drawing.Width = width;
        drawing.Height = height;

        // Use the same values for the embedded picture.
        var pic = drawing.Picture!;
        pic.Width = width;
        pic.Height = height;
            
        return drawing;
    }
}