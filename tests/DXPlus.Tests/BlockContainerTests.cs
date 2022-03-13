using System.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class BlockContainerTests
    {
        [Fact]
        public void ParagraphCanFindDocumentOwner()
        {
            using var doc = Document.Create();

            var p = doc.Add("This is a paragraph.");
            Assert.NotNull(p.Container);
            Assert.Equal(doc as BlockContainer, p.Container);
        }

        [Fact]
        public void CanEnumerateAllBlocks()
        {
            var doc = Document.Create();

            doc.Add("This is paragraph #1.");
            doc.Add(new Table(3, 2));
            doc.Add("This is paragraph #2.");

            var blocks = doc.Paragraphs.ToList();

            Assert.Equal(2, blocks.Count);

            Assert.IsType<Paragraph>(blocks[0]);
            Assert.IsType<Paragraph>(blocks[1]);
            Assert.NotNull(blocks[0].Table);

            Assert.Equal(3, doc.Blocks.Count());
            Assert.IsType<Paragraph>(doc.Blocks.First());
            Assert.IsType<Table>(doc.Blocks.ElementAt(1));
            Assert.IsType<Paragraph>(doc.Blocks.Last());
        }

        [Fact]
        public void RemoveParagraphAtIndex0()
        {
            var doc = Document.Create();

            doc.Add("This is paragraph #1.");
            doc.Add("This is paragraph #2.");
            doc.Add("This is paragraph #3.");
            doc.Add("This is paragraph #4.");
            Assert.Equal(4, doc.Paragraphs.Count());

            doc.RemoveAt(0);
            Assert.Equal(3, doc.Paragraphs.Count());
            Assert.Equal("This is paragraph #2.", doc.Paragraphs.First().Text);
        }

        [Fact]
        public void RemoveParagraphFromMiddle()
        {
            var doc = Document.Create();

            doc.Add("This is paragraph #1.");
            doc.Add("This is paragraph #2.");
            doc.Add("This is paragraph #3.");
            doc.Add("This is paragraph #4.");
            Assert.Equal(4, doc.Paragraphs.Count());

            doc.RemoveAt(2);

            var paragraphs = doc.Paragraphs.ToList();

            Assert.Equal(3, paragraphs.Count);
            Assert.Equal("This is paragraph #1.", paragraphs[0].Text);
            Assert.Equal("This is paragraph #2.", paragraphs[1].Text);
            Assert.Equal("This is paragraph #4.", paragraphs[2].Text);
        }

        [Fact]
        public void RemoveParagraphFromEnd()
        {
            var doc = Document.Create();

            doc.Add("This is paragraph #1.");
            doc.Add("This is paragraph #2.");
            doc.Add("This is paragraph #3.");
            doc.Add("This is paragraph #4.");
            Assert.Equal(4, doc.Paragraphs.Count());

            doc.RemoveAt(3);
            Assert.Equal(3, doc.Paragraphs.Count());
            Assert.Equal("This is paragraph #3.", doc.Paragraphs.ElementAt(2).Text);
        }
        
    }
}