using System.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class SectionTests
    {
        [Fact]
        public void NewDocHasSingleSection()
        {
            var doc = Document.Create();
            Assert.Single(doc.Sections);
        }

        [Fact]
        public void DefaultSectionIncludesAllParagraphs()
        {
            var doc = Document.Create();

            var p1 = doc.AddParagraph("This is the first paragraph");
            var p2 = doc.AddParagraph("This is the second paragraph");
            var p3 = doc.AddParagraph("This is the third paragraph");

            Assert.Single(doc.Sections);
            var section = doc.Sections.Single();
            Assert.Equal(3, section.Paragraphs.Count);
            Assert.Contains(p1, section.Paragraphs);
            Assert.Equal(p2.Text, section.Paragraphs[1].Text);
            Assert.Equal(p3.Text, section.Paragraphs[2].Text);
        }

        [Fact]
        public void SectionIncludesAllParagraphs()
        {
            var doc = Document.Load("MultipleSectionsTest.docx");

            Assert.Equal(3, doc.Sections.Count);

            var s1 = doc.Sections[0];
            Assert.Equal(6, s1.Paragraphs.Count);
            Assert.NotNull(s1.Paragraphs.Last().Xml.Element(Name.ParagraphProperties)?.Element(Name.SectionProperties));
            Assert.True(s1.Footers.Even.Exists);
            Assert.True(s1.Footers.Default.Exists);
            Assert.False(s1.Footers.First.Exists);
            Assert.Equal(2, s1.Footers.Count());
            Assert.Empty(s1.Headers);

            var s2 = doc.Sections[1];
            Assert.Equal(5, s2.Paragraphs.Count);
            Assert.Equal(144, s2.Properties.LeftMargin);
            Assert.Equal(72, s2.Properties.TopMargin);
            Assert.Empty(s2.Headers);
            Assert.Empty(s2.Footers);

            var s3 = doc.Sections[2];
            Assert.Equal(5, s3.Paragraphs.Count);
            Assert.Equal(SectionBreakType.EvenPage, s3.Properties.Type);
            Assert.Empty(s3.Headers);
            Assert.Empty(s3.Footers);
        }
    }
}
