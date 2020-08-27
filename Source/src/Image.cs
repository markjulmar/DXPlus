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
        private readonly IDocument document;

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

            return ((DocX)document).Package.GetPart(new Uri(full, UriKind.Relative)).GetStream(mode, access);
        }

        /// <summary>
        /// Returns the id of this Image.
        /// </summary>
        public string Id { get; }

        internal Image(IDocument document, PackageRelationship packageRelationship)
        {
            this.document = document;
            this.PackageRelationship = packageRelationship;
            Id = packageRelationship.Id;
        }

        /// <summary>
        /// Create a new picture and insert it into a paragraph.
        /// </summary>
        /// <returns>New picture</returns>
        public Picture CreatePicture()
        {
            return Paragraph.CreatePicture(document as DocX, Id, string.Empty, string.Empty);
        }

        /// <summary>
        /// Create a new picture with specific dimensions and insert it into a paragraph.
        /// </summary>
        /// <returns>New picture</returns>
        public Picture CreatePicture(int height, int width)
        {
            var picture = Paragraph.CreatePicture(document as DocX, Id, string.Empty, string.Empty);
            picture.Height = height;
            picture.Width = width;
            return picture;
        }

        ///<summary>
        /// Returns the name of the image file.
        ///</summary>
        public string FileName => Path.GetFileName(PackageRelationship.TargetUri.ToString());
    }
}