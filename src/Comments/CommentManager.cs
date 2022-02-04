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
        private PackagePart commentsPackagePart;

        private static readonly XName CommentStart = Namespace.Main + "commentRangeStart";
        private static readonly XName CommentEnd = Namespace.Main + "commentRangeEnd";
        private static readonly XName CommentReference = Namespace.Main + "commentReference";

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="documentOwner">Owning document</param>
        public CommentManager(IDocument documentOwner) 
            : base(documentOwner, null, null)
        {
        }

        /// <summary>
        /// Retrieve the comments package part
        /// </summary>
        internal PackagePart CommentsPackagePart
        {
            get => commentsPackagePart;
            set
            {
                commentsPackagePart = value;
                commentsDoc = PackagePart.Load();
                Xml = commentsDoc?.Root;
            }
        }

        /// <summary>
        /// Loads the people.xml from the document
        /// </summary>
        internal PackagePart PeoplePackagePart
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
                : Xml.Elements(Namespace.Main + "comment")
                     .Select(e => new Comment(Document, CommentsPackagePart, e));

        /// <summary>
        /// The comment package part
        /// </summary>
        internal override PackagePart PackagePart => commentsPackagePart;

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

            if (commentsPackagePart == null)
            {
                commentsPackagePart = Document.Package.CreatePart(Relations.Comments.Uri, Relations.Comments.ContentType, CompressionOption.Maximum);
                var template = Resource.CommentsDocument();
                commentsPackagePart.Save(template);
                Document.PackagePart.CreateRelationship(commentsPackagePart.Uri, TargetMode.Internal, Relations.Comments.RelType);
                commentsDoc = commentsPackagePart.Load();
                Xml = commentsDoc.Root;
            }

            var person = peopleDoc.Root?.Elements(Namespace.W2012 + "person")
                                          .FirstOrDefault(p => p.Attribute(Namespace.W2012 + "author")?.Value == authorName);
            if (person == null)
            {
                AddPersonEntity(authorName);
            }

            var comment = new Comment(Document, commentsPackagePart, authorName, dt, authorInitials)
            {
                Id = Comments.Any() ? Comments.Max(c => c.Id) + 1 : 1,
            };

            // Insert into the DOM
            Xml?.Add(comment.Xml);

            return comment;
        }

        /// <summary>
        /// Add a new person to the comment authors document.
        /// TODO: it doesn't support any auth providers.
        /// </summary>
        /// <param name="authorName">Author name to add</param>
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
        /// Retrieve the comments tied to a specific paragraph in the document.
        /// </summary>
        /// <param name="owner">Paragraph owner</param>
        /// <returns>Enumerable of comment ranges</returns>
        public IEnumerable<CommentRange> GetCommentsForParagraph(Paragraph owner)
        {
            // Look for a w:commentReference tag in the paragraph.
            return owner.Xml.Descendants(CommentReference)
                            .Select(commentRef => ParseComment(owner, commentRef))
                            .Where(comment => comment != null);
        }

        /// <summary>
        /// Parses out the range for a given comment.
        /// </summary>
        /// <param name="owner"></param>
        /// <param name="commentReference"></param>
        /// <returns></returns>
        private CommentRange ParseComment(Paragraph owner, XElement commentReference)
        {
            var commentId = HelperFunctions.GetId(commentReference) ?? -1;
            if (commentId == -1) 
                return null;

            // Find the commentRangeStart for this id.
            var commentStart = Document.Xml.Descendants(CommentStart)
                .Single(cs => HelperFunctions.GetId(cs) == commentId);
            var commentEnd = Document.Xml.Descendants(CommentEnd)
                .Single(cs => HelperFunctions.GetId(cs) == commentId);

            // Now find all the runs between these two tags.
            var nodes = Document.Xml.Descendants()
                .SkipWhile(node => node != commentStart)
                .TakeWhile(node => node != commentEnd)
                .Where(e => e.Name == Name.Run && !e.Descendants(CommentReference).Any())
                .ToList();

            var srXml = nodes.FirstOrDefault();
            var erXml = nodes.LastOrDefault();

            var startRun = srXml != null ? Document.FindRunByElement(srXml) : null;
            var endRun = erXml == null ? null : erXml == srXml ? startRun : Document.FindRunByElement(erXml);

            return new CommentRange(owner, startRun, endRun, Comments.Single(c => c.Id == commentId));
        }

        /// <summary>
        /// Attach a comment to a run
        /// </summary>
        /// <param name="comment"></param>
        /// <param name="runStart"></param>
        /// <param name="runEnd"></param>
        public static void Attach(Comment comment, Run runStart, Run runEnd)
        {
            runStart.Xml.AddBeforeSelf(
                new XElement(CommentStart,
                    new XAttribute(Name.Id, comment.Id)));

            var endNode = new XElement(CommentEnd,
                    new XAttribute(Name.Id, comment.Id));

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
