using System;
using System.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class ParagraphTests
    {
        private const string Filename = "test.docx";

        [Fact]
        public void AlignmentGetAndSetAreAligned()
        {
            using var doc = DocX.Create(Filename);
            doc.InsertParagraph("This is a test.").Align(Alignment.Center);
            Assert.Equal(Alignment.Center, doc.Paragraphs[0].Alignment);
        }

        [Fact]
        public void DefaultAlignmentIsLeft()
        {
            using var doc = DocX.Create(Filename);
            doc.InsertParagraph("This is a test.");
            Assert.Equal(Alignment.Left, doc.Paragraphs[0].Alignment);
        }

        [Fact]
        public void DirectionGetAndSetAreAligned()
        {
            using var doc = DocX.Create(Filename);
            var p = doc.InsertParagraph("This is a test.");
            p.Direction = Direction.RightToLeft;
            Assert.Equal(Direction.RightToLeft, doc.Paragraphs[0].Direction);
        }

        [Fact]
        public void DefaultDirectionIsLeftToRight()
        {
            using var doc = DocX.Create(Filename);
            doc.InsertParagraph("This is a test.");
            Assert.Equal(Direction.LeftToRight, doc.Paragraphs[0].Direction);
        }

        [Fact]
        public void AddDocumentProperty()
        {
            using var doc = DocX.Create(Filename);

            var props = new[]
            {
                new CustomProperty("intProperty", 100),
                new CustomProperty("stringProperty", "100"),
                new CustomProperty("doubleProperty", 100.0),
                new CustomProperty("dateProperty", new DateTime(2010, 1, 1)),
                new CustomProperty("boolProperty", true)
            };

            var p = doc.InsertParagraph();
            props.ToList().ForEach(prop => p.AddDocumentProperty(prop));

            Assert.Equal(props.Length, p.DocumentProperties.Count);
            for (var index = 0; index < p.DocumentProperties.Count; index++)
            {
                var prop = p.DocumentProperties[index];
                Assert.Equal(props[index].Name, prop.Name);
            }
        }

        [Fact]
        public void CheckHeading1()
        {
            using var doc = DocX.Create(Filename);
            doc.InsertParagraph("This is a test.").Heading(HeadingType.Heading1);
            Assert.Equal("Heading1", doc.Paragraphs[0].StyleName);
        }

        [Fact]
        public void CheckHyperlinksInDoc()
        {
            var microsoftUrl = new Uri("http://www.microsoft.com");
            using var doc = DocX.Create(Filename);
            doc.InsertParagraph("Test");
            doc.InsertParagraph()
                .AppendLine("This line contains a ")
                .Append(new Hyperlink("link", microsoftUrl))
                .Append(".");

            Assert.Single(doc.Hyperlinks);
            Assert.Equal("link", doc.Hyperlinks[0].Text);
            Assert.Equal(microsoftUrl, doc.Hyperlinks[0].Uri);
        }

        [Fact]
        public void CheckHeaders()
        {
            using var doc = DocX.Create(Filename);
            doc.InsertParagraph("Test");
            doc.InsertSectionPageBreak();
            doc.InsertParagraph("Second page.");

            Assert.NotNull(doc.Headers);
            Assert.Null(doc.Headers.Even);
            Assert.Null(doc.Headers.Odd);
            Assert.Null(doc.Headers.First);

            doc.AddHeaders();
            Assert.NotNull(doc.Headers.Even);
            Assert.NotNull(doc.Headers.Odd);
            Assert.NotNull(doc.Headers.First);

            doc.Headers.First.InsertParagraph("This is the first page header");
            doc.SaveAs(@"/users/mark/Desktop/test.docx");
        }
    }
}
