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
        public string FileName => PackageRelationship != null ? Path.GetFileName(PackageRelationship.TargetUri.ToString()) : string.Empty;

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
        /// Create a new picture and insert it into a paragraph.
        /// </summary>
        /// <returns>New picture</returns>
        public Picture CreatePicture(string name = null, string description = null) =>
            document.CreatePicture(Id, name??string.Empty, description??string.Empty);

        /// <summary>
        /// Create a new picture with specific dimensions and insert it into a paragraph.
        /// </summary>
        /// <returns>New picture</returns>
        public Picture CreatePicture(int width, int height) =>
            CreatePicture(string.Empty, string.Empty, width, height);

        /// <summary>
        /// Create a new picture with specific dimensions and insert it into a paragraph.
        /// </summary>
        /// <returns>New picture</returns>
        public Picture CreatePicture(string name, string description, int width, int height)
        {
            var picture = CreatePicture(name, description);
            picture.Width = width;
            picture.Height = height;
            return picture;
        }
    }
}