using System.Xml.XPath;
using Xunit;

namespace DXPlus.Tests
{
    public class BookmarkTests
    {
        [Fact]
        public void AddBookmarkInsertsXml()
        {
            DocX doc = (DocX) Document.Create();

            doc.AddParagraph("This is a test paragraph.");
            var p = doc
                .AddParagraph("This is a second paragraph with a book mark - ")
                .AppendBookmark("bookmark1")
                .Append(" added into the text.");

            Assert.True(p.BookmarkExists("bookmark1"));
            Assert.Single(doc.Xml.RemoveNamespaces().XPathSelectElements("//bookmarkStart"));
        }
    }
}