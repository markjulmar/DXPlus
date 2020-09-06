using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using Xunit;

namespace DXPlus.Tests
{
    public class ParagraphTests
    {
        private const string Filename = "test.docx";

        [Fact]
        public void AddCustomDocumentProperty()
        {
            using var doc = Document.Create();

            doc.AddCustomProperty("intProperty", 100);
            doc.AddCustomProperty("stringProperty", "100");
            doc.AddCustomProperty("doubleProperty", 100.5);
            doc.AddCustomProperty("dateProperty", new DateTime(2010, 1, 1));
            doc.AddCustomProperty("boolProperty", true);

            Assert.Equal(5, doc.CustomProperties.Count);

            var p = doc.AddParagraph();

            p.AddCustomPropertyField("doubleProperty");
            Assert.Single(p.DocumentProperties);

            var prop = p.DocumentProperties.Single();
            Assert.Equal("doubleProperty", prop.Name);
            Assert.Equal("100.5", prop.Value);

            p.AppendLine();

            p.AddCustomPropertyField("dateProperty");
            Assert.Equal(2, p.DocumentProperties.Count());

            prop = p.DocumentProperties.Skip(1).Single();
            Assert.Equal("dateProperty", prop.Name);
            Assert.Equal(new DateTime(2010, 1, 1).ToString(), prop.Value);
        }

        [Fact]
        public void AddDocumentProperty()
        {
            const string text = "The title.";
            using var doc = Document.Create();

            doc.SetPropertyValue(DocumentPropertyName.Title, text);
            Assert.Equal(text, doc.DocumentProperties[DocumentPropertyName.Title]);

            var p = doc.AddParagraph();
            p.AddDocumentPropertyField(DocumentPropertyName.Title);

            Assert.Single(p.DocumentProperties);

            var prop = p.DocumentProperties.Single();
            Assert.Equal("TITLE", prop.Name);
            Assert.Equal(text, prop.Value);
        }


        [Fact]
        public void CheckHeading1()
        {
            using var doc = Document.Create(Filename);
            doc.AddParagraph("This is a test.").Style(HeadingType.Heading1);
            Assert.Equal("Heading1", doc.Paragraphs[0].Properties.StyleName);
        }

        [Fact]
        public void CheckHyperlinksInDoc()
        {
            var microsoftUrl = new Uri("http://www.microsoft.com");
            using var doc = Document.Create(Filename);
            doc.AddParagraph("Test");
            doc.AddParagraph()
                .AppendLine("This line contains a ")
                .Append(new Hyperlink("link", microsoftUrl))
                .Append(".");

            Assert.Single(doc.Hyperlinks);
            Assert.Equal("link", doc.Hyperlinks.First().Text);
            Assert.Equal(microsoftUrl, doc.Hyperlinks.First().Uri);
        }

        [Fact]
        public void FirstHeaderAddsElements()
        {
            using var doc = (DocX) Document.Create();

            var section = doc.Sections.Single();

            Assert.NotNull(section.Headers);
            Assert.False(section.Headers.Even.Exists);
            Assert.False(section.Headers.Default.Exists);
            Assert.False(section.Headers.First.Exists);
            Assert.False(section.Properties.DifferentFirstPage);

            section.Headers.First.Add().SetText("Page Header 1");
            Assert.False(section.Headers.Even.Exists);
            Assert.False(section.Headers.Default.Exists);
            Assert.True(section.Headers.First.Exists);
            Assert.True(section.Properties.DifferentFirstPage);

            Assert.Single(doc.Xml.RemoveNamespaces().XPathSelectElements("//sectPr/headerReference"));
            Assert.Empty(doc.Xml.RemoveNamespaces().XPathSelectElements("//sectPr/footerReference"));
        }

        [Fact]
        public void FirstHeaderRemovesElements()
        {
            using var doc = (DocX)Document.Create();
            var section = doc.Sections.Single();
            Assert.NotNull(section.Headers);

            section.Headers.First.Add().SetText("Page Header 1");
            Assert.False(section.Headers.Even.Exists);
            Assert.False(section.Headers.Default.Exists);
            Assert.True(section.Headers.First.Exists);
            Assert.True(section.Properties.DifferentFirstPage);

            section.Headers.First.Remove();
            Assert.False(section.Headers.Even.Exists);
            Assert.False(section.Headers.Default.Exists);
            Assert.False(section.Headers.First.Exists);
            Assert.False(section.Properties.DifferentFirstPage);

            Assert.Empty(doc.Xml.RemoveNamespaces().XPathSelectElements("//sectPr/headerReference"));
            Assert.Empty(doc.Xml.RemoveNamespaces().XPathSelectElements("//sectPr/footerReference"));
        }

        [Fact]
        public void SecondHeaderIncrementsCorrectly()
        {
            using var doc = (DocX)Document.Create();
            var section = doc.Sections.Single();

            Assert.NotNull(section.Headers);
            Assert.False(section.Headers.Even.Exists);
            Assert.False(section.Headers.Default.Exists);
            Assert.False(section.Headers.First.Exists);

            section.Headers.First.Add();
            Assert.Equal("/word/header1.xml", section.Headers.First.Uri.OriginalString);
            section.Headers.Even.Add();
            Assert.Equal("/word/header2.xml", section.Headers.Even.Uri.OriginalString);

            Assert.True(section.Headers.Even.Exists);
            Assert.False(section.Headers.Default.Exists);
            Assert.True(section.Headers.First.Exists);

            Assert.Equal(2, doc.Xml.RemoveNamespaces().XPathSelectElements("//sectPr/headerReference").Count());
            Assert.Empty(doc.Xml.RemoveNamespaces().XPathSelectElements("//sectPr/footerReference"));
        }

        [Fact]
        public void WithFormattingAddsNewProperties()
        {
            var p = new Paragraph();
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));

            p.WithFormatting(new Formatting {Bold = true});
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/b"));
        }

        [Fact]
        public void WithFormattingNullRemovesParagraphProperties()
        {
            var p = new Paragraph().WithFormatting(new Formatting());
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));

            p.WithFormatting(null);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));
        }

        [Fact]
        public void WithFormattingReplacesProperties()
        {
            var p = new Paragraph().WithFormatting(new Formatting {Bold = true});
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/rPr/b"));

            p.WithFormatting(new Formatting {Italic = true});
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/i"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/b"));
        }

        [Fact]
        public void WithFormattingAffectsRun()
        {
            var p = new Paragraph("This is a test").WithFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr/b"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/rPr"));
        }

        [Fact]
        public void WithFormattingNullRemovesRunProperties()
        {
            var p = new Paragraph("This is a test").WithFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));

            p.WithFormatting(null);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));
        }

        [Fact]
        public void WithFormattingAffectsLastRun()
        {
            var p = new Paragraph("This is a test")
                .AppendLine("With a second line")
                .AppendLine("and a final line")
                .WithFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr/b"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/rPr"));

            var runs = p.Runs.Reverse().ToList();
            Assert.Equal(5, runs.Count);
            Assert.True(runs[1].Properties.Bold); // skip LF
            Assert.Equal("and a final line", runs[1].Text);
        }

        [Fact]
        public void StyleAddsRemovesElement()
        {
            var p = new Paragraph();
            Assert.Equal("Normal", p.Properties.StyleName);

            p.Properties.StyleName = "Body";
            Assert.Equal("Body", p.Properties.StyleName);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/pStyle"));

            p.Properties.StyleName = null;
            Assert.Equal("Normal", p.Properties.StyleName);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/pStyle"));

            p.Properties.StyleName = "";
            Assert.Equal("Normal", p.Properties.StyleName);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/pStyle"));
        }

        [Fact]
        public void SetTextReplacesContents()
        {
            var p = new Paragraph();
            p.InsertText("This is a test. ");
            p.AppendLine("Will it work?");
            Assert.Equal("This is a test. Will it work?\n", p.Text);

            p.SetText("of the emergency broadcast system.");
            Assert.Equal("of the emergency broadcast system.", p.Text);
            Assert.Equal("of the emergency broadcast system.", p.Xml.RemoveNamespaces().Value);
        }

        [Fact]
        public void CheckPackagePartAssignmentForParagraphs()
        {
            using DocX doc = (DocX) Document.Create();
            Assert.NotNull(doc.PackagePart);

            Paragraph p = new Paragraph("Test");
            Assert.Null(p.PackagePart);

            doc.AddParagraph(p);
            Assert.NotNull(p.PackagePart);
            Assert.Equal(doc.PackagePart, p.PackagePart);
        }

        [Fact]
        public void InsertParagraphAtZeroAddsToBeginning()
        {
            using var doc = Document.Create(Filename);
            var firstParagraph = doc.AddParagraph("First paragraph");
            Assert.Single(doc.Paragraphs);
            Assert.Equal("First paragraph", firstParagraph.Text);

            var secondParagraph = doc.AddParagraph("Another paragraph");
            Assert.Equal(2, doc.Paragraphs.Count);
            Assert.Equal("Another paragraph", secondParagraph.Text);

            var p = doc.InsertParagraph(0, " Inserted Text ");
            Assert.Equal(3, doc.Paragraphs.Count);

            Assert.Equal(" Inserted Text ", doc.Paragraphs[0].Text);
        }

        [Fact]
        public void InsertParagraphInMiddleSplitsParagraph()
        {
            using var doc = Document.Create(Filename);
            var firstParagraph = doc.AddParagraph("First paragraph");
            Assert.Single(doc.Paragraphs);
            Assert.Equal("First paragraph", firstParagraph.Text);

            var p = doc.InsertParagraph(5, " Inserted Text ");
            Assert.Equal(3, doc.Paragraphs.Count);

            Assert.Equal("First", doc.Paragraphs[0].Text);
            Assert.Equal(" Inserted Text ", doc.Paragraphs[1].Text);
            Assert.Equal(" paragraph", doc.Paragraphs[2].Text);
        }

        [Fact]
        public void RemoveTextEditsParagraph()
        {
            const string text = "This is a paragraph in a document where we are looking to remove some text.";
            using var doc = Document.Create(Filename);
            var p = doc.AddParagraph(text);
            Assert.Single(doc.Paragraphs);
            Assert.Equal(text, p.Text);

            p.RemoveText(10, 10);
            Assert.Equal("This is a in a document where we are looking to remove some text.", p.Text);
        }

        [Fact]
        public void RemoveTextRemovesEmptyParagraph()
        {
            const string text = "Test";
            using var doc = Document.Create();
            var p = doc.AddParagraph(text);
            Assert.Single(doc.Paragraphs);
            Assert.Equal(text, p.Text);

            p.RemoveText(0);
            Assert.Empty(doc.Paragraphs);
        }

        [Fact]
        public void RemoveTextWithMultipleRunsSpansRun()
        {
            using var doc = Document.Create(Filename);
            var p = doc.AddParagraph("This")
                .Append(" is ").WithFormatting(new Formatting() { Bold = true })
                .Append("a ").WithFormatting(new Formatting() { Bold = true })
                .Append("test.");

            Assert.Equal("This is a test.", p.Text);

            p.RemoveText(5, 3);
            Assert.Equal("This a test.", p.Text);
        }

        [Fact]
        public void RemoveTextWithInsertRemovesText()
        {
            var e = new XElement(Name.Paragraph,
                new XElement(Name.Run,
                    new XElement(Name.Text, "Some ")),
                new XElement(Namespace.Main + "ins",
                    new XElement(Name.Run,
                        new XElement(Name.Text, "text goes "))),
                new XElement(Name.Run,
                    new XElement(Name.Text, "here.")));

            var p = new Paragraph(null, e, 0);
            Assert.Equal("Some text goes here.", p.Text);
            Assert.Equal(0, p.StartIndex);

            p.RemoveText(5, 5);
            Assert.Equal("Some goes here.", p.Text);
        }

        [Fact]
        public void RemoveTextWithInsertRemovesPartialText()
        {
            var e = new XElement(Name.Paragraph,
                new XElement(Name.Run,
                    new XElement(Name.Text, "Some ")),
                new XElement(Namespace.Main + "ins",
                    new XElement(Name.Run,
                        new XElement(Name.Text, "text goes "))),
                new XElement(Name.Run,
                    new XElement(Name.Text, "here.")));

            var p = new Paragraph(null, e, 0);
            Assert.Equal("Some text goes here.", p.Text);
            Assert.Equal(0, p.StartIndex);

            p.RemoveText(4, 6);
            Assert.Equal("Somegoes here.", p.Text);
        }

        [Fact]
        public void ParagraphPropertiesAlwaysInsertFirst()
        {
            Paragraph p = new Paragraph();

            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("pPr"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));

            p.WithFormatting(new Formatting());
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("pPr"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));

            var pPr = p.Xml.Element(Name.ParagraphProperties);
            Assert.Equal(p.Xml, pPr.Parent);
            Assert.Equal(p.Xml.Descendants().First(), pPr);
        }

        [Fact]
        public void ParagraphRunPropertiesAlwaysInsertFirst()
        {
            Paragraph p = new Paragraph();

            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("pPr"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));

            p.SetText("This is a test");

            p.WithFormatting(new Formatting());
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("pPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));

            var lastRun = p.Xml.Elements(Name.Run).Last();
            Assert.NotNull(lastRun);

            var rPr = p.Xml.XPathSelectElements("//w:r/w:rPr", Namespace.NamespaceManager()).Single();
            Assert.NotNull(rPr);

            Assert.Equal("r", rPr.Parent.Name.LocalName);
            Assert.Equal(lastRun, rPr.Parent);
        }

        [Fact]
        public void FormattingChangesPreviousTextRun()
        {
            var p = new Paragraph();

            p.Append("This is a test").WithFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r"));

            var lastRun = p.Xml.Elements(Name.Run).Last();
            var properties = lastRun.Element(Name.RunProperties);
            Assert.Single(properties.Elements(Name.Bold));
        }

        [Fact]
        public void FormattingChangesSkipLineBreaks()
        {
            var p = new Paragraph();

            p.AppendLine("This is a test").WithFormatting(new Formatting { Bold = true });
            Assert.Equal(2, p.Xml.RemoveNamespaces().XPathSelectElements("//r").Count());

            var textRun = p.Xml.Elements(Name.Run).First();
            var properties = textRun.Element(Name.RunProperties);
            Assert.Single(properties.Elements(Name.Bold));
        }

        [Fact]
        public void TwoParagraphsFromSameSectionAreTheSame()
        {
            var doc = Document.Create();

            var p1 = doc.AddParagraph("This is a test");
            var p2 = doc.Paragraphs[0];

            Assert.Equal(p1,p2);  // value equality
            Assert.False(p1==p2); // not reference equality
        }

        [Fact]
        public void EquationRunIsReturnedInRunCollection()
        {
            var p = new Paragraph("Test paragraph");
            Assert.Single(p.Runs);

            p.AppendEquation("1 + 2 = 3");
            Assert.Equal(2, p.Runs.Count());
        }
    }
}
