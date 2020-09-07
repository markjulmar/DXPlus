using Xunit;

namespace DXPlus.Tests
{
    public class BlockContainerTests
    {
        [Fact]
        public void RemoveParagraphAtIndex0()
        {
            var doc = Document.Create();

            doc.AddParagraph("This is paragraph #1.");
            doc.AddParagraph("This is paragraph #2.");
            doc.AddParagraph("This is paragraph #3.");
            doc.AddParagraph("This is paragraph #4.");
            Assert.Equal(4, doc.Paragraphs.Count);

            doc.RemoveParagraph(0);
            Assert.Equal(3, doc.Paragraphs.Count);
            Assert.Equal("This is paragraph #2.", doc.Paragraphs[0].Text);
        }

        [Fact]
        public void RemoveParagraphFromMiddle()
        {
            var doc = Document.Create();

            doc.AddParagraph("This is paragraph #1.");
            doc.AddParagraph("This is paragraph #2.");
            doc.AddParagraph("This is paragraph #3.");
            doc.AddParagraph("This is paragraph #4.");
            Assert.Equal(4, doc.Paragraphs.Count);

            doc.RemoveParagraph(2);
            Assert.Equal(3, doc.Paragraphs.Count);
            Assert.Equal("This is paragraph #1.", doc.Paragraphs[0].Text);
            Assert.Equal("This is paragraph #2.", doc.Paragraphs[1].Text);
            Assert.Equal("This is paragraph #4.", doc.Paragraphs[2].Text);
        }

        [Fact]
        public void RemoveParagraphFromEnd()
        {
            var doc = Document.Create();

            doc.AddParagraph("This is paragraph #1.");
            doc.AddParagraph("This is paragraph #2.");
            doc.AddParagraph("This is paragraph #3.");
            doc.AddParagraph("This is paragraph #4.");
            Assert.Equal(4, doc.Paragraphs.Count);

            doc.RemoveParagraph(3);
            Assert.Equal(3, doc.Paragraphs.Count);
            Assert.Equal("This is paragraph #3.", doc.Paragraphs[2].Text);
        }
        
    }
}