using System;
using System.IO;
using System.IO.Packaging;

namespace DXPlus
{
    /// <summary>
    /// Represents an Image embedded in a document.
    /// </summary>
    public class Image
    {
        private readonly Document document;

        /// <summary>
        /// Associated package relationship
        /// </summary>
        internal PackageRelationship PackageRelationship { get; }

        /// <summary>
        /// Get the stream for the picture
        /// </summary>
        /// <param name="mode">File mode</param>
        /// <param name="access">Access type</param>
        /// <returns>Open stream</returns>
        public Stream OpenStream()
        {
            string temp = PackageRelationship.SourceUri.OriginalString;
            string start = temp.Remove(temp.LastIndexOf('/'));
            string end = PackageRelationship.TargetUri.OriginalString;
            string full = start + "/" + end;

            // Return a readonly stream to the image
            return document.Package.GetPart(new Uri(full, UriKind.Relative))
                           .GetStream(FileMode.Open, FileAccess.Read);
        }

        /// <summary>
        /// Returns the id of this Image.
        /// </summary>
        public string Id { get; }

        ///<summary>
        /// Returns the name of the image file.
        ///</summary>
        public string FileName => PackageRelationship != null ? Path.GetFileName(PackageRelationship.TargetUri.ToString()) : string.Empty;

        /// <summary>
        /// Returns the image type (extension) such as .png, .jpg, .svg, etc.
        /// </summary>
        public string ImageType => Path.GetExtension(FileName)?.ToLower();

        /// <summary>
        /// Internal constructor to create an image from a package relationship.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="packageRelationship"></param>
        internal Image(IDocument document, PackageRelationship packageRelationship)
        {
            this.document = (Document) document;
            this.PackageRelationship = packageRelationship;
            Id = packageRelationship.Id;
        }

        /// <summary>
        /// Internal constructor to create an image from an unowned document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="packageRelationship"></param>
        /// <param name="id"></param>
        internal Image(IDocument document, PackageRelationship packageRelationship, string id)
        {
            this.document = (Document)document;
            this.PackageRelationship = packageRelationship;
            this.Id = id;
        }

        /// <summary>
        /// Create a new picture which can be added to a paragraph from this image.
        /// </summary>
        /// <param name="name">Name of the image</param>
        /// <param name="description">Description of the image</param>
        /// <returns>New picture</returns>
        public Drawing CreatePicture(string name = null, string description = null) =>
            document.CreateDrawingWithEmbeddedPicture(Id, name??string.Empty, description??string.Empty);

        /// <summary>
        /// Create a new picture which can be added to a paragraph from this image.
        /// </summary>
        /// <param name="width">Width of the picture</param>
        /// <param name="height">Height of the picture</param>
        /// <returns>New picture</returns>
        public Drawing CreatePicture(int width, int height) =>
            CreatePicture(string.Empty, string.Empty, width, height);

        /// <summary>
        /// Create a new picture with specific dimensions and insert it into a paragraph.
        /// </summary>
        /// <param name="name">Name of the picture</param>
        /// <param name="description">Description of the picture</param>
        /// <param name="height">Rendered height</param>
        /// <param name="width">Rendered width</param>
        /// <returns>New picture</returns>
        public Drawing CreatePicture(string name, string description, int width, int height)
        {
            // Create the drawing.
            var drawing = CreatePicture(name, description);
            
            // Set the extent on the drawing
            drawing.Width = width;
            drawing.Height = height;

            // Use the same values for the embedded picture.
            var pic = drawing.Picture;
            pic.Width = width;
            pic.Height = height;
            
            return drawing;
        }
    }
}