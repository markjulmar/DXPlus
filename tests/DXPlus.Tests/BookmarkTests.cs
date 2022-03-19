using System;
using System.Linq;
using System.Xml.XPath;
using Xunit;

namespace DXPlus.Tests
{
    public class BookmarkTests
    {
        [Fact]
        public void AddBookmarkInsertsXml()
        {
            Document doc = (Document) Document.Create();

            doc.AddParagraph("This is a test paragraph.");
            var p = doc
                .AddParagraph("This is a second paragraph with a book mark - ")
                .SetBookmark("bookmark1")
                .AddText(" added into the text.");

            Assert.True(p.BookmarkExists("bookmark1"));
            Assert.Single(doc.Xml.RemoveNamespaces().XPathSelectElements("//bookmarkStart"));

            var bookmark = p.Bookmarks[0];

            Assert.Equal(bookmark, p.Bookmarks["bookmark1"]);
            Assert.Equal("bookmark1",bookmark.Name);
            Assert.Equal(p, bookmark.Paragraph);
            Assert.Equal(p.Runs.First().Text, bookmark.Text);

            Assert.Throws<ArgumentException>(() => p.SetBookmark("bookmark1"));
        }

        [Fact]
        public void AddBookMarkCanSetOnRuns()
        {
            Document doc = (Document)Document.Create();

            var p = doc
                .AddParagraph("This is a test paragraph.")
                .AddText(" With lots of text.")
                .AddText(" Added over time.")
                .AddText(" And a final sentence.");

            var runs = p.Runs.ToList();

            p.SetBookmark("bookmark1", runs[1], runs[2]);
            var bookmark = p.Bookmarks[0];

            Assert.Equal(" With lots of text. Added over time.", bookmark.Text);
        }
    }
}