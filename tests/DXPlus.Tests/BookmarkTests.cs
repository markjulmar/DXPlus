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

            doc.Add("This is a test paragraph.");
            var p = doc
                .Add("This is a second paragraph with a book mark - ")
                .AddBookmark("bookmark1")
                .AddText(" added into the text.");

            Assert.True(p.BookmarkExists("bookmark1"));
            Assert.Single(doc.Xml.RemoveNamespaces().XPathSelectElements("//bookmarkStart"));

            var bookmark = p.Bookmarks[0];
            bookmark.SetText("HI!");

            Assert.Equal(bookmark, p.Bookmarks["bookmark1"]);
            Assert.Equal("bookmark1",bookmark.Name);
            Assert.Equal(p, bookmark.Paragraph);
            Assert.Equal("HI!", bookmark.Text);
            Assert.Equal("This is a second paragraph with a book mark - HI! added into the text.", p.Text);

            Assert.Throws<ArgumentException>(() => p.SetBookmark("bookmark1"));
        }

        [Fact]
        public void AddBookMarkCanSetOnRuns()
        {
            Document doc = (Document)Document.Create();

            var p = doc
                .Add("This is a test paragraph.")
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