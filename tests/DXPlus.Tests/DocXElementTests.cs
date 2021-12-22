using System;
using System.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class DocXElementTests
    {
        [Fact]
        public void CheckParagraphEqualityOperators()
        {
            using var doc = Document.Create();
            var p1 = doc.AddParagraph("Test");
            var p2 = doc.Paragraphs.First();

            Assert.False(object.ReferenceEquals(p1, p2));
            Assert.False(p1 == p2); // should be ref test
            Assert.False(object.Equals(p1, p2));
            Assert.True(p1.Equals(p2));
            Assert.Equal(p1, p2);
        }

        [Fact]
        public void CheckHLEqualityOperators()
        {
            using var doc = Document.Create();
            var hl1 = new Hyperlink("test", new Uri("https://google.com"));
            var p1 = doc.AddParagraph("This is a ").Append(hl1);

            var hl2 = p1.Hyperlinks.First();

            Assert.False(object.ReferenceEquals(hl1,hl2));
            Assert.False(hl1 == hl2); // should be ref test
            Assert.False(object.Equals(hl1, hl2));
            Assert.True(hl1.Equals(hl2));
        }
    }
}
