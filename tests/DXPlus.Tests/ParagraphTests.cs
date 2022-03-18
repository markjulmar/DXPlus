using System;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using DXPlus.Internal;
using Xunit;

namespace DXPlus.Tests
{
    public class ParagraphTests
    {
        private const string Filename = "test.docx";

        [Fact]
        public void CannotAddParagraphToOrphanParagraph()
        {
            var p = new Paragraph();
            Assert.Throws<InvalidOperationException>(() => p.AddParagraph());
        }

        [Fact]
        public void OnceParentedCanAddParagraphs()
        {
            using var doc = Document.Create();
            var p = new Paragraph();
            doc.Add(p);
            p.AddParagraph();
            Assert.Equal(2, doc.Paragraphs.Count());
        }

        [Fact]
        public void AddToDocSetsStartAndEndIndex()
        {
            using var doc = Document.Create();

            var p = new Paragraph("Test");
            Assert.Null(p.StartIndex);

            doc.Add(p);
            Assert.Equal(0, p.StartIndex);

            p = new Paragraph("Second paragraph");
            Assert.Null(p.StartIndex);

            doc.Add(p);
            Assert.Equal("Test".Length, p.StartIndex);
            Assert.Equal("Test\nSecond paragraph", doc.Text);
        }

        [Fact]
        public void AddCustomDocumentProperty()
        {
            using var doc = Document.Create();

            doc.CustomProperties.Add("intProperty", 100);
            doc.CustomProperties.Add("stringProperty", "100");
            doc.CustomProperties.Add("doubleProperty", 100.5);
            doc.CustomProperties.Add("dateProperty", new DateTime(2010, 1, 1));
            doc.CustomProperties.Add("boolProperty", true);

            Assert.Equal(5, doc.CustomProperties.Count);

            var p = doc.AddParagraph();

            p.AddCustomPropertyField("doubleProperty");
            Assert.Single(p.Fields);

            var prop = p.Fields.Single();
            Assert.Equal("doubleProperty", prop.Name);
            Assert.Equal("100.5", prop.Value);

            p.Newline();

            p.AddCustomPropertyField("dateProperty");
            Assert.Equal(2, p.Fields.Count());

            prop = p.Fields.Skip(1).Single();
            Assert.Equal("dateProperty", prop.Name);
            Assert.Equal(new DateTime(2010, 1, 1), DateTime.Parse(prop.Value));
        }

        [Fact]
        public void AddDocumentProperty()
        {
            const string text = "The title.";
            using var doc = Document.Create();

            doc.Properties.Title = text;
            Assert.Equal(text, doc.Properties.Title);

            var p = doc.AddParagraph();
            p.AddDocumentPropertyField(DocumentPropertyName.Title);

            Assert.Single(p.Fields);

            var prop = p.Fields.Single();
            Assert.Equal("TITLE", prop.Name);
            Assert.Equal(text, prop.Value);
        }

        [Fact]
        public void SingleParaDocReturnsNoNextPrevParagraph()
        {
            using var doc = Document.Create();
            var p = doc.Add("1");

            Assert.Null(p.NextParagraph);
            Assert.Null(p.PreviousParagraph);
        }

        [Fact]
        public void FirstParagraphReturnsNullPrevParagraph()
        {
            using var doc = Document.Create();
            var p = doc.Add("1");
            doc.Add("2");
            doc.Add("3");

            Assert.Null(p.PreviousParagraph);
            Assert.NotNull(p.NextParagraph);
        }

        [Fact]
        public void LastParagraphReturnsNullNextParagraph()
        {
            using var doc = Document.Create();
            doc.Add("1");
            doc.Add("2");
            var p = doc.Add("3");

            Assert.NotNull(p.PreviousParagraph);
            Assert.Null(p.NextParagraph);
        }

        [Fact]
        public void PreviousParagraphReturnsProperObject()
        {
            using var doc = Document.Create();
            doc.Add("1");
            doc.Add("2");
            var p1 = doc.Add("3");
            doc.Add("4");
            var p2 = doc.Add("5");
            var p3 = doc.Add("6");

            Assert.Equal("2", p1.PreviousParagraph.Text);
            Assert.Equal("4", p2.PreviousParagraph.Text);
            Assert.Equal("5", p3.PreviousParagraph.Text);
        }

        [Fact]
        public void NextParagraphReturnsProperObject()
        {
            using var doc = Document.Create();
            var p1 = doc.Add("1");
            doc.Add("2");
            var p2 = doc.Add("3");
            doc.Add("4");
            var p3 = doc.Add("5");
            doc.Add("6");

            Assert.Equal("2", p1.NextParagraph.Text);
            Assert.Equal("4", p2.NextParagraph.Text);
            Assert.Equal("6", p3.NextParagraph.Text);
        }

        [Fact]
        public void CheckHeading1()
        {
            using var doc = Document.Create(Filename);
            doc.Add("This is a test.").Style(HeadingType.Heading1);
            Assert.Equal("Heading1", doc.Paragraphs.First().Properties.StyleName);
        }

        [Fact]
        public void SingleSectionOwnsAllParagraphs()
        {
            var doc = Document.Create();
            for (int i = 0; i < 10; i++)
            {
                doc.Add($"Paragraph #{i + 1}");
            }

            Assert.Equal(10, doc.Paragraphs.Count());
            Assert.Single(doc.Sections);

            var bodySection = doc.Sections.Single();
            foreach (var p in doc.Paragraphs)
            {
                Assert.True(bodySection.Equals(p.Section));
            }
        }

        [Fact]
        public void MultiSectionContainsSpecificParagraphs()
        {
            var doc = Document.Create();
            int i;
            for (i = 0; i < 5; i++)
            {
                doc.Add($"Paragraph #{i + 1}");
            }

            doc.AddSection(); // All paragraphs above.

            for (; i < 10; i++)
            {
                doc.Add($"Paragraph #{i + 1}");
            }

            Assert.Equal(11, doc.Paragraphs.Count()); // add one for section
            Assert.Equal(2, doc.Sections.Count());

            var firstSection = doc.Sections.First();
            var bodySection = doc.Sections.Last();

            i = 0;
            foreach (var p in doc.Paragraphs.Where(p => !p.IsSectionParagraph))
            {
                Assert.True(((i<5) ? firstSection : bodySection).Equals(p.Section));
                i++;
            }

        }

        [Fact]
        public void CheckHyperlinksInDoc()
        {
            var microsoftUrl = new Uri("http://www.microsoft.com");
            using var doc = Document.Create(Filename);
            doc.Add("Test");
            doc.AddParagraph()
                .Add("This line contains a ")
                .Add(new Hyperlink("link", microsoftUrl))
                .Add(".");

            Assert.Single(doc.Hyperlinks);
            Assert.Equal("link", doc.Hyperlinks.First().Text);
            Assert.Equal(microsoftUrl, doc.Hyperlinks.First().Uri);
        }

        [Fact]
        public void CheckHyperlinksInUnownedParagraph()
        {
            var microsoftUrl = new Uri("http://www.microsoft.com");

            var paragraph = new Paragraph()
                .Add("This line contains a ")
                .Add(new Hyperlink("link", microsoftUrl))
                .Add(".");

            Assert.Single(paragraph.Hyperlinks);
            Assert.Equal("link", paragraph.Hyperlinks.First().Text);
            Assert.Empty(paragraph.Hyperlinks.First().Id);
            Assert.Equal(microsoftUrl, paragraph.Hyperlinks.First().Uri);
        }

        [Fact]
        public void HyperlinkGetsIdWhenAddedToDocument()
        {
            var microsoftUrl = new Uri("http://www.microsoft.com");

            var paragraph = new Paragraph()
                .Add("This line contains a ")
                .Add(new Hyperlink("link", microsoftUrl))
                .Add(".");

            Assert.Single(paragraph.Hyperlinks);
            Assert.Empty(paragraph.Hyperlinks.First().Id);

            var doc = Document.Create();
            doc.Add((Paragraph) paragraph);
            Assert.NotEmpty(paragraph.Hyperlinks.First().Id);
        }

        [Fact]
        public void FirstHeaderAddsElements()
        {
            using var doc = (Document) Document.Create();

            var section = doc.Sections.Single();

            Assert.NotNull(section.Headers);
            Assert.False(section.Headers.Even.Exists);
            Assert.False(section.Headers.Default.Exists);
            Assert.False(section.Headers.First.Exists);
            Assert.False(section.Properties.DifferentFirstPage);

            section.Headers.First.Paragraphs.First().Text = "Page Header 1";
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
            using var doc = (Document)Document.Create();
            var section = doc.Sections.Single();
            Assert.NotNull(section.Headers);

            section.Headers.First.MainParagraph.Text = "Page Header 1";
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
            using var doc = (Document)Document.Create();
            var section = doc.Sections.Single();

            Assert.NotNull(section.Headers);
            Assert.False(section.Headers.Even.Exists);
            Assert.False(section.Headers.Default.Exists);
            Assert.False(section.Headers.First.Exists);

            _ = section.Headers.First.MainParagraph;
            Assert.Equal("/word/header1.xml", section.Headers.First.Uri.OriginalString);
            _ = section.Headers.Even.MainParagraph;
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

            p.MergeFormatting(new Formatting {Bold = true});
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/b"));
        }

        [Fact]
        public void ClearFormattingNullRemovesParagraphProperties()
        {
            var p = new Paragraph().MergeFormatting(new Formatting());
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));

            p.ClearFormatting();
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));
        }

        [Fact]
        public void WithFormattingMergesProperties()
        {
            var p = new Paragraph().MergeFormatting(new Formatting {Bold = true});
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/rPr/b"));

            p.MergeFormatting(new Formatting {Italic = true});
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/i"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/b"));
        }

        [Fact]
        public void WithFormattingAffectsRun()
        {
            var p = new Paragraph("This is a test").MergeFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr/b"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/rPr"));
        }

        [Fact]
        public void ClearFormattingNullRemovesRunProperties()
        {
            var p = new Paragraph("This is a test").MergeFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));

            Assert.Throws<ArgumentNullException>(() => p.MergeFormatting(null));

            p.ClearFormatting();
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));
        }

        [Fact]
        public void WithFormattingAffectsLastRun()
        {
            var p = new Paragraph("This is a test")
                .Add("With a second line")
                .Add("and a final line")
                .MergeFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr/b"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/rPr"));

            var runs = p.Runs.Reverse().ToList();
            Assert.Equal(3, runs.Count);
            Assert.True(runs[0].Properties.Bold); // skip LF
            Assert.Equal("and a final line", runs[0].Text);
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
            p.Add("This is a test. ");
            p.Add("Will it work?");
            p.Newline();
            Assert.Equal("This is a test. Will it work?\n", p.Text);

            p.Text = "of the emergency broadcast system.";
            Assert.Equal("of the emergency broadcast system.", p.Text);
            Assert.Equal("of the emergency broadcast system.", p.Xml.RemoveNamespaces().Value);
        }

        [Fact]
        public void CheckPackagePartAssignmentForParagraphs()
        {
            using Document doc = (Document) Document.Create();
            Assert.NotNull(doc.PackagePart);

            Paragraph p = new Paragraph("Test");
            Assert.Throws<InvalidOperationException>(() => p.PackagePart);

            doc.Add(p);
            Assert.NotNull(p.PackagePart);
            Assert.Equal(doc.PackagePart, p.PackagePart);
        }

        /* NullContainer allows this to work.
        [Fact]
        public void AddPageBreakToOrphanedParagraphThrowsException()
        {
            var paragraph = new FirstParagraph("Test");
            Assert.Throws<InvalidOperationException>(() => paragraph.AddPageBreak());
        }
        */

        [Fact]
        public void AppendAddsParagraph()
        {
            using var doc = Document.Create();
            var firstParagraph = doc.Add("First paragraph");
            var secondParagraph = doc.Add("Another paragraph");

            var p = firstParagraph.InsertAfter(new Paragraph("Injected paragraph"));
            Assert.Equal(3, doc.Paragraphs.Count());

            var ps = doc.Paragraphs.ToList();
            Assert.Equal(p.Id, ps[1].Id);
        }

        [Fact]
        public void InsertBeforeAddsParagraph()
        {
            using var doc = Document.Create();
            var firstParagraph = doc.Add("First paragraph");
            var secondParagraph = doc.Add("Another paragraph");

            var p = secondParagraph.InsertBefore(new Paragraph("Injected paragraph"));
            Assert.Equal(3, doc.Paragraphs.Count());

            var ps = doc.Paragraphs.ToList();
            Assert.Equal(p.Id, ps[1].Id);
        }

        [Fact]
        public void InsertBeforeFirstParagraphAddsNewIntro()
        {
            using var doc = Document.Create();
            doc.Add("Introduction").Style(HeadingType.Heading1);
            
            doc.Add("Some text goes here.")
               .Add("With more text");

            Assert.Equal(2, doc.Paragraphs.Count());

            var p = doc.Paragraphs.First()
                .InsertBefore(new Paragraph("Title").Style(HeadingType.Title));
            Assert.Equal(3, doc.Paragraphs.Count());

            var ps = doc.Paragraphs.ToList();
            Assert.Equal(p.Id, ps[0].Id);
            Assert.Equal(HeadingType.Heading1.ToString(), ps[1].Properties.StyleName);
        }

        [Fact]
        public void AppendToLastParagraphAddsSummary()
        {
            using var doc = Document.Create();
            var firstParagraph = doc.Add("First paragraph");
            var secondParagraph = doc.Add("Another paragraph");

            var p = secondParagraph.InsertAfter(new Paragraph("Injected paragraph"));
            Assert.Equal(3, doc.Paragraphs.Count());

            var ps = doc.Paragraphs.ToList();
            Assert.Equal(p.Id, ps.Last().Id);
        }

        [Fact]
        public void InsertParagraphAtZeroAddsToBeginning()
        {
            using var doc = Document.Create(Filename);
            var firstParagraph = doc.Add("First paragraph");
            Assert.Single(doc.Paragraphs);
            Assert.Equal("First paragraph", firstParagraph.Text);

            var secondParagraph = doc.Add("Another paragraph");
            Assert.Equal(2, doc.Paragraphs.Count());
            Assert.Equal("Another paragraph", secondParagraph.Text);

            var p = doc.Insert(0, " Inserted Text ");
            Assert.Equal(3, doc.Paragraphs.Count());

            Assert.Equal(" Inserted Text ", doc.Paragraphs.First().Text);
        }
        
        [Fact]
        public void InsertParagraphInMiddleSplitsParagraph()
        {
            using var doc = Document.Create(Filename);
            var firstParagraph = doc.Add("First paragraph");
            Assert.Single(doc.Paragraphs);
            Assert.Equal("First paragraph", firstParagraph.Text);

            var p = doc.Insert(5, " Inserted Text ");
            Assert.Equal(3, doc.Paragraphs.Count());

            Assert.Equal("First", doc.Paragraphs.First().Text);
            Assert.Equal(" Inserted Text ", doc.Paragraphs.ElementAt(1).Text);
            Assert.Equal(" paragraph", doc.Paragraphs.ElementAt(2).Text);
        }

        [Fact]
        public void RemoveTextEditsParagraph()
        {
            const string text = "This is a paragraph in a document where we are looking to remove some text.";
            using var doc = Document.Create(Filename);
            var p = doc.Add(text);
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
            var p = doc.Add(text);
            Assert.Single(doc.Paragraphs);
            Assert.Equal(text, p.Text);

            p.RemoveText(0);
            Assert.Empty(doc.Paragraphs);
        }

        [Fact]
        public void RemoveTextWithMultipleRunsSpansRun()
        {
            using var doc = Document.Create(Filename);
            var p = doc.Add("This")
                .Add(" is ").MergeFormatting(new Formatting() { Bold = true })
                .Add("a ").MergeFormatting(new Formatting() { Bold = true })
                .Add("test.");

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
                new XElement(Namespace.Main + RunTextType.InsertMarker,
                    new XElement(Name.Run,
                        new XElement(Name.Text, "text goes "))),
                new XElement(Name.Run,
                    new XElement(Name.Text, "here.")));

            var p = new Paragraph(null, null, e, null);
            Assert.Equal("Some text goes here.", p.Text);
            Assert.Null(p.StartIndex);

            p.RemoveText(5, 5);
            Assert.Equal("Some goes here.", p.Text);
        }

        [Fact]
        public void RemoveTextWithInsertRemovesPartialText()
        {
            var e = new XElement(Name.Paragraph,
                new XElement(Name.Run,
                    new XElement(Name.Text, "Some ")),
                new XElement(Namespace.Main + RunTextType.InsertMarker,
                    new XElement(Name.Run,
                        new XElement(Name.Text, "text goes "))),
                new XElement(Name.Run,
                    new XElement(Name.Text, "here.")));

            var p = new Paragraph(null, null, e, 0);
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

            p.MergeFormatting(new Formatting());
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("pPr"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//r/rPr"));

            var pPr = p.Xml.Element(Name.ParagraphProperties);
            Assert.Equal(p.Xml, pPr.Parent);
            Assert.Equal(p.Xml.Descendants().First(), pPr);
        }

        [Fact]
        public void ParagraphRunPropertiesAlwaysInsertFirst()
        {
            var p = new Paragraph();

            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("pPr"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr"));

            p.Text = "This is a test";

            p.MergeFormatting(new Formatting());
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

            p.Add("This is a test").MergeFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r"));

            var lastRun = p.Xml.Elements(Name.Run).Last();
            var properties = lastRun.Element(Name.RunProperties);
            Assert.Single(properties.Elements(Name.Bold));
        }

        [Fact]
        public void FormattingChangesSkipLineBreaks()
        {
            var p = new Paragraph();
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//r"));

            p.Add("This is a test").MergeFormatting(new Formatting { Bold = true });
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r"));

            var textRun = p.Xml.Elements(Name.Run).First();
            var properties = textRun.Element(Name.RunProperties);
            Assert.NotNull(properties);
            Assert.Single(properties.Elements(Name.Bold));
        }

        [Fact]
        public void TwoParagraphsFromSameSectionAreTheSame()
        {
            var doc = Document.Create();

            var p1 = doc.Add("This is a test");
            var p2 = doc.Paragraphs.First();

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

        [Fact]
        public void UnownedParagraphInsertsImage()
        {
            var document = Document.Create();
            var image = document.CreateImage("1022.jpg");
            var drawing = image.CreatePicture(150, 150);
            var pic = drawing.Picture;

            var paragraph = new Paragraph();
            paragraph.Add(drawing);

            Assert.Single(paragraph.Pictures);
            Assert.Same(pic.Xml, paragraph.Pictures.Single().Xml);
        }

        [Fact]
        public void AssigningDocOwnerUpdatesImage()
        {
            var document = Document.Create();
            var image = document.CreateImage("1022.jpg");
            var picture = image.CreatePicture(150, 150);

            var paragraph = new Paragraph();
            paragraph.Add(picture);

            Assert.Single(paragraph.Pictures);
            Assert.Null(paragraph.Pictures[0].SafePackagePart);

            document.Add(paragraph);
            Assert.NotNull(paragraph.Pictures[0].PackagePart);
        }

        [Fact]
        public void AddMultipleParagraphsWithAddRange()
        {
            var document = Document.Create();
            document.AddRange(new [] { "Starting text", "More text", "Another paragraph", "Last paragraph" });
            Assert.Equal(4, document.Paragraphs.Count());
        }

        [Fact]
        public void AddCaptionToImage()
        {
            string text = ": This is a picture.";

            var document = Document.Create();

            document.Add("Starting paragraph.");

            var paragraph = document.AddParagraph();
            var image = document.CreateImage("1022.jpg");
            var picture = image.CreatePicture(150, 150);
            
            paragraph.Add(picture);
            picture.AddCaption(text);

            document.Add("Ending paragraph");

            Assert.Equal("Figure 1" + text, picture.GetCaption());
            Assert.Throws<ArgumentException>(() => picture.AddCaption(text));

            var picture2 = image.CreatePicture(200, 200);
            document.AddParagraph().Add(picture2);
            picture2.AddCaption("Another picture");

            Assert.Equal("Figure 2 Another picture", picture2.GetCaption());
        }

        [Fact]
        public void FindReplaceReturnsFalseWhenNotFound()
        {
            var p = new Paragraph("This is a test paragraph");
            Assert.False(p.FindReplace("tst", "Test"));

            p = new Paragraph();
            Assert.False(p.FindReplace("test", null));
        }

        [Fact]
        public void FindReplaceReplaceCharacter()
        {
            var p = new Paragraph("This is a test paragraph");
            Assert.True(p.FindReplace("test", "Test"));
            Assert.Equal("This is a Test paragraph", p.Text);
            Assert.Throws<ArgumentNullException>(() => p.FindReplace(null, null));
        }

        [Fact]
        public void FindReplaceCanRemoveWords()
        {
            var p = new Paragraph("This is a test paragraph");
            Assert.True(p.FindReplace("test ", null));
            Assert.Equal("This is a paragraph", p.Text);

            Assert.True(p.FindReplace("paragraph", null));
            Assert.Equal("This is a ", p.Text);

            Assert.True(p.FindReplace("This is a ", null));
            Assert.Equal("", p.Text);
        }

        [Fact]
        public void TextSkipsDeletedText()
        {
            string xml =
                @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:r>
				        <w:t xml:space=""preserve"">This is a test with </w:t>
			        </w:r>
			        <w:del>
				        <w:r>
					        <w:delText xml:space=""preserve"">deleted </w:delText>
				        </w:r>
			        </w:del>
			        <w:ins>
				        <w:r>
					        <w:t xml:space=""preserve"">inserted </w:t>
				        </w:r>
			        </w:ins>
			        <w:r>
				        <w:t>text.</w:t>
			        </w:r>
		        </w:p>";

            var p = new Paragraph(null, null, XElement.Parse(xml), null);

            Assert.Equal("This is a test with inserted text.", p.Text);
        }
    }
}
