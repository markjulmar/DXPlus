using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;
using DXPlus.Resources;

namespace DXPlus.Comments
{
    internal class CommentManager : DocXElement
    {
        private XDocument commentsDoc;
        private XDocument peopleDoc;
        private PackagePart peoplePackagePart;

        private static readonly XName CommentStart = Namespace.Main + "commentRangeStart";
        private static readonly XName CommentEnd = Namespace.Main + "commentRangeEnd";

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="documentOwner">Owning document</param>
        public CommentManager(IDocument documentOwner) 
            : base(documentOwner, null)
        {
        }

        /// <summary>
        /// Loads the people.xml from the document
        /// </summary>
        public PackagePart PeoplePackagePart
        {
            get => peoplePackagePart;
            set
            {
                peoplePackagePart = value ?? throw new ArgumentNullException(nameof(value));
                peopleDoc = peoplePackagePart.Load();
            }
        }

        /// <summary>
        /// Returns all comments in the document
        /// </summary>
        public IEnumerable<Comment> Comments 
            => Xml == null 
                ? Enumerable.Empty<Comment>() 
                : Xml.Elements(Namespace.Main + "comment").Select(e => new Comment(Document, e));

        /// <summary>
        /// The comment package part
        /// </summary>
        public PackagePart CommentPackagePart
        {
            get => PackagePart;
            set
            {
                PackagePart = value;
                commentsDoc = PackagePart.Load();
                Xml = commentsDoc?.Root;
            }
        }

        /// <summary>
        /// Creates a new comment with a blank paragraph
        /// </summary>
        /// <param name="authorName">Author name</param>
        /// <param name="dt">Optional data</param>
        /// <param name="authorInitials">Optional initials</param>
        /// <returns>Created comment</returns>
        public Comment CreateComment(string authorName, DateTime? dt = null, string authorInitials = null)
        {
            if (peoplePackagePart == null)
            {
                peoplePackagePart = Document.Package.CreatePart(Relations.People.Uri, Relations.People.ContentType, CompressionOption.Maximum);
                var template = Resource.PeopleDocument();
                peoplePackagePart.Save(template);
                Document.PackagePart.CreateRelationship(peoplePackagePart.Uri, TargetMode.Internal, Relations.People.RelType);
                peopleDoc = peoplePackagePart.Load();
            }

            if (PackagePart == null)
            {
                PackagePart = Document.Package.CreatePart(Relations.Comments.Uri, Relations.Comments.ContentType, CompressionOption.Maximum);
                var template = Resource.CommentsDocument();
                PackagePart.Save(template);
                Document.PackagePart.CreateRelationship(PackagePart.Uri, TargetMode.Internal, Relations.Comments.RelType);
                commentsDoc = PackagePart.Load();
                Xml = commentsDoc.Root;
            }

            var person = peopleDoc.Root?.Elements(Namespace.W2012 + "person")
                                          .FirstOrDefault(p => p.Attribute(Namespace.W2012 + "author")?.Value == authorName);
            if (person == null)
            {
                AddPersonEntity(authorName);
            }

            var comment = new Comment(Document, authorName, dt, authorInitials)
            {
                Id = Comments.Any() ? Comments.Max(c => c.Id) + 1 : 1,
                PackagePart = PackagePart
            };

            // Insert into the DOM
            Xml?.Add(comment.Xml);

            return comment;
        }

        private void AddPersonEntity(string authorName)
        {
            peopleDoc.Root?.Add(
                new XElement(Namespace.W2012 + "person",
                    new XAttribute(Namespace.W2012 + "author", authorName),
                    new XElement(Namespace.W2012 + "presenceInfo",
                        new XAttribute(Namespace.W2012 + "providerId", "None"),
                        new XAttribute(Namespace.W2012 + "userId", authorName))));
        }

        /// <summary>
        /// Attach a comment to a run
        /// </summary>
        /// <param name="comment"></param>
        /// <param name="runStart"></param>
        /// <param name="runEnd"></param>
        public void Attach(Comment comment, Run runStart, Run runEnd)
        {
            runStart.Xml.AddBeforeSelf(
                new XElement(CommentStart, new XAttribute(Namespace.Main + "id", comment.Id)));
            var endNode = new XElement(CommentEnd, new XAttribute(Namespace.Main + "id", comment.Id));
            runEnd.Xml.AddAfterSelf(endNode);
            endNode.AddAfterSelf(XElement.Parse(
                        $@"<w:r xmlns:w=""{Namespace.Main}"">
                            <w:rPr>
                            <w:rStyle w:val=""CommentReference""/>
                            <w:color w:val=""auto""/>
                            </w:rPr>
                            <w:commentReference w:id=""{comment.Id}""/>
                        </w:r>"));
        }

        /// <summary>
        /// Save changes to the comments
        /// </summary>
        public void Save()
        {
            PackagePart?.Save(commentsDoc);
            peoplePackagePart?.Save(peopleDoc);
        }
    }
}
