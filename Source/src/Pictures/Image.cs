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
        private readonly DocX document;

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
        public Stream GetStream(FileMode mode, FileAccess access)
        {
            string temp = PackageRelationship.SourceUri.OriginalString;
            string start = temp.Remove(temp.LastIndexOf('/'));
            string end = PackageRelationship.TargetUri.OriginalString;
            string full = start + "/" + end;

            return document.Package.GetPart(new Uri(full, UriKind.Relative)).GetStream(mode, access);
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
        /// Internal constructor to create an image from a package relationship.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="packageRelationship"></param>
        internal Image(IDocument document, PackageRelationship packageRelationship)
        {
            this.document = (DocX) document;
            this.PackageRelationship = packageRelationship;
            Id = packageRelationship.Id;
        }

        /// <summary>
        /// Create a new picture and insert it into a paragraph.
        /// </summary>
        /// <returns>New picture</returns>
        public Picture CreatePicture(string name = null, string description = null) =>
            document.CreatePicture(Id, name??string.Empty, description??string.Empty);

        /// <summary>
        /// Create a new picture with specific dimensions and insert it into a paragraph.
        /// </summary>
        /// <returns>New picture</returns>
        public Picture CreatePicture(int height, int width) =>
            CreatePicture(string.Empty, string.Empty, height, width);

        /// <summary>
        /// Create a new picture with specific dimensions and insert it into a paragraph.
        /// </summary>
        /// <returns>New picture</returns>
        public Picture CreatePicture(string name, string description, int height, int width)
        {
            var picture = CreatePicture(name, description);
            picture.Height = height;
            picture.Width = width;
            return picture;
        }
    }
}