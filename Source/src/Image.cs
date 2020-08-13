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
        internal PackageRelationship pr;

        public Stream GetStream(FileMode mode, FileAccess access)
        {
            string temp = pr.SourceUri.OriginalString;
            string start = temp.Remove(temp.LastIndexOf('/'));
            string end = pr.TargetUri.OriginalString;
            string full = start + "/" + end;

            return document.package.GetPart(new Uri(full, UriKind.Relative)).GetStream(mode, access);
        }

        /// <summary>
        /// Returns the id of this Image.
        /// </summary>
        public string Id { get; }

        internal Image(DocX document, PackageRelationship pr)
        {
            this.document = document;
            this.pr = pr;
            this.Id = pr.Id;
        }

        /// <summary>
        /// Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
        /// </summary>
        public Picture CreatePicture()
        {
            return Paragraph.CreatePicture(document, Id, string.Empty, string.Empty);
        }

        public Picture CreatePicture(int height, int width)
        {
            Picture picture = Paragraph.CreatePicture(document, Id, string.Empty, string.Empty);
            picture.Height = height;
            picture.Width = width;
            return picture;
        }

        ///<summary>
        /// Returns the name of the image file.
        ///</summary>
        public string FileName => Path.GetFileName(pr.TargetUri.ToString());
    }
}