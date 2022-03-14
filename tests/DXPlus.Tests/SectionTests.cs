using System.Linq;
using DXPlus.Internal;
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
        public void CanFindContainerFromHeaderParagraph()
        {
            using var doc = Document.Create();
            var mainSection = doc.Sections.First();
            var header = mainSection.Headers.Default;

            var p1 = header.Paragraphs.Single();
            p1.Text = "This is some text - ";
            p1.AddPageNumber(PageNumberFormat.Normal);

            doc.Add("P1");
            doc.AddPageBreak();
            doc.Add("P2");

            var p2 = header.Paragraphs.First();
            Assert.Equal(p1, p2);

            Assert.NotNull(p2.Container);
            Assert.Equal(header, p2.Container);
        }

        [Fact]
        public void CanFindContainerFromFooterParagraph()
        {
            using var doc = Document.Create();
            var mainSection = doc.Sections.First();

            var footer = mainSection.Footers.Default;
            var p = footer.Paragraphs.Single().AddParagraph("New paragraph");

            Assert.NotNull(p.Container);
            Assert.Equal(footer, p.Container);
        }

        [Fact]
        public void DefaultSectionIncludesAllParagraphs()
        {
            var doc = Document.Create();

            var p1 = doc.Add("This is the first paragraph");
            var p2 = doc.Add("This is the second paragraph");
            var p3 = doc.Add("This is the third paragraph");

            Assert.Single(doc.Sections);
            var section = doc.Sections.Single();
            Assert.Equal(3, section.Paragraphs.Count());
            Assert.Contains(p1, section.Paragraphs);
            Assert.Equal(p2.Text, section.Paragraphs.ElementAt(1).Text);
            Assert.Equal(p3.Text, section.Paragraphs.ElementAt(2).Text);
        }

        [Fact]
        public void SectionIncludesAllParagraphs()
        {
            var doc = Document.Load("MultipleSectionsTest.docx");

            var sections = doc.Sections.ToList();

            Assert.Equal(3, sections.Count);

            var s1 = sections[0];
            Assert.Equal(6, s1.Paragraphs.Count());
            Assert.NotNull(s1.Paragraphs.Last().Xml.Element(Name.ParagraphProperties)?.Element(Name.SectionProperties));
            Assert.True(s1.Footers.Even.Exists);
            Assert.True(s1.Footers.Default.Exists);
            Assert.False(s1.Footers.First.Exists);
            Assert.Equal(2, s1.Footers.Count());
            Assert.Empty(s1.Headers);

            var s2 = sections[1];
            Assert.Equal(5, s2.Paragraphs.Count());
            Assert.Equal(2880, s2.Properties.LeftMargin);
            Assert.Equal(1440, s2.Properties.TopMargin);
            Assert.Empty(s2.Headers);
            Assert.Empty(s2.Footers);

            var s3 = sections[2];
            Assert.Equal(5, s3.Paragraphs.Count());
            Assert.Equal(SectionBreakType.EvenPage, s3.Properties.Type);
            Assert.Empty(s3.Headers);
            Assert.Empty(s3.Footers);
        }
    }
}
