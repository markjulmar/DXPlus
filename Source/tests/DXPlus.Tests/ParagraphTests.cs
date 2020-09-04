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
        public void AlignmentGetAndSetAreAligned()
        {
            using var doc = Document.Create(Filename);
            doc.AddParagraph("This is a test.").Alignment(Alignment.Center);
            Assert.Equal(Alignment.Center, doc.Paragraphs[0].Alignment);
        }

        [Fact]
        public void DefaultAlignmentIsLeft()
        {
            using var doc = Document.Create(Filename);
            doc.AddParagraph("This is a test.");
            Assert.Equal(Alignment.Left, doc.Paragraphs[0].Alignment);
        }

        [Fact]
        public void DirectionGetAndSetAreAligned()
        {
            var p = new Paragraph("This is a test.") {Direction = Direction.RightToLeft};
            Assert.Equal(Direction.RightToLeft, p.Direction);
        }

        [Fact]
        public void DefaultDirectionIsLeftToRight()
        {
            using var doc = Document.Create(Filename);
            doc.AddParagraph("This is a test.");
            Assert.Equal(Direction.LeftToRight, doc.Paragraphs[0].Direction);
        }

        [Fact]
        public void AddDocumentProperty()
        {
            using var doc = Document.Create(Filename);

            var props = new[]
            {
                new CustomProperty("intProperty", 100),
                new CustomProperty("stringProperty", "100"),
                new CustomProperty("doubleProperty", 100.0),
                new CustomProperty("dateProperty", new DateTime(2010, 1, 1)),
                new CustomProperty("boolProperty", true)
            };

            var p = doc.AddParagraph();
            props.ToList().ForEach(prop => p.AddDocumentProperty(prop));

            Assert.Equal(props.Length, p.DocumentProperties.Count());
            int index = 0;
            foreach(var prop in p.DocumentProperties)
            {
                Assert.Equal(props[index++].Name, prop.Name);
            }
        }

        [Fact]
        public void CheckHeading1()
        {
            using var doc = Document.Create(Filename);
            doc.AddParagraph("This is a test.").Heading(HeadingType.Heading1);
            Assert.Equal("Heading1", doc.Paragraphs[0].StyleName);
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
        public void BoldAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.False(p.Bold);

            p.Bold = true;
            Assert.True(p.Bold);

            Assert.NotNull(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/b"));

            p.Bold = false;
            Assert.False(p.Bold);
            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/b"));
        }

        [Fact]
        public void ItalicAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.False(p.Italic);

            p.Italic = true;
            Assert.True(p.Italic);

            Assert.NotNull(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/i"));

            p.Italic = false;
            Assert.False(p.Italic);
            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/i"));
        }

        [Fact]
        public void CapsStyleAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Equal(CapsStyle.None, p.CapsStyle);

            p.CapsStyle = CapsStyle.Caps;
            Assert.Equal(CapsStyle.Caps, p.CapsStyle);

            Assert.NotNull(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/caps"));
            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/smallCaps"));

            p.CapsStyle = CapsStyle.SmallCaps;
            Assert.Equal(CapsStyle.SmallCaps, p.CapsStyle);

            Assert.NotNull(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/smallCaps"));
            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/caps"));

            p.CapsStyle = CapsStyle.None;
            Assert.Equal(CapsStyle.None, p.CapsStyle);

            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/smallCaps"));
            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/caps"));
        }

        [Fact]
        public void ColorAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Equal(Color.Empty, p.Color);

            p.Color = Color.Red;
            Assert.NotStrictEqual(Color.Red, p.Color);
            Assert.NotNull(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/color"));

            p.Color = Color.Empty;
            Assert.Equal(Color.Empty, p.Color);
            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/color"));
        }

        [Fact]
        public void CultureAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Null(p.Culture);

            var spanish = CultureInfo.GetCultureInfo("es-BR");

            p.Culture = spanish;
            Assert.Equal(spanish, p.Culture);
            Assert.NotNull(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/lang"));

            p.Culture = null;
            Assert.Null(p.Culture);
            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/lang"));

            p.Culture();
            Assert.Equal(CultureInfo.CurrentCulture, p.Culture);
            Assert.NotNull(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/lang"));
        }

        [Fact]
        public void FontAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Null(p.Font);

            var ff = new FontFamily("Times New Roman");

            p.Font = ff;
            Assert.Equal(ff.Name, p.Font.Name);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/rFonts"));

            ff = new FontFamily("Wingdings");
            p.Font = ff;
            Assert.Equal(ff.Name, p.Font.Name);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/rFonts"));

            p.Font = null;
            Assert.Null(p.Font);
            Assert.Null(p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/rFonts"));
        }

        [Fact]
        public void FontSizeAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Null(p.FontSize);

            p.FontSize = 32;
            Assert.Equal(32, p.FontSize);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/sz"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/szCs"));
            Assert.Equal("64", p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/sz").Attribute("val")?.Value);

            p.FontSize = 22;
            Assert.Equal(22, p.FontSize);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/sz"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/szCs"));
            Assert.Equal("44", p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/sz").Attribute("val")?.Value);

            p.FontSize = null;
            Assert.Null(p.FontSize);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/sz"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/szCs"));
        }

        [Fact]
        public void HideAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.False(p.IsHidden);

            p.IsHidden = true;
            Assert.True(p.IsHidden);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vanish"));

            // Make sure we don't dup the tag
            p.IsHidden = true;
            Assert.True(p.IsHidden);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vanish"));

            p.IsHidden = false;
            Assert.False(p.IsHidden);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vanish"));
        }

        [Fact]
        public void HighlightAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Equal(Highlight.None, p.Highlight);

            p.Highlight = Highlight.Yellow;
            Assert.Equal(Highlight.Yellow, p.Highlight);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/highlight"));

            // Make sure we don't dup the tag
            p.Highlight = Highlight.Green;
            Assert.Equal(Highlight.Green, p.Highlight);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/highlight"));

            p.Highlight = Highlight.None;
            Assert.Equal(Highlight.None, p.Highlight);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/highlight"));
        }

        [Fact]
        public void EffectAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Equal(Effect.None, p.Effect);

            p.Effect = Effect.Emboss;
            Assert.Equal(Effect.Emboss, p.Effect);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/emboss"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/shadow"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outline"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outlineShadow"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/imprint"));

            // Make sure we don't dup the tag
            p.Effect = Effect.Shadow;
            Assert.Equal(Effect.Shadow, p.Effect);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/emboss"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/shadow"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outline"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outlineShadow"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/imprint"));

            p.Effect = Effect.None;
            Assert.Equal(Effect.None, p.Effect);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/emboss"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/shadow"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outline"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outlineShadow"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/imprint"));

            p.Effect = Effect.OutlineShadow;
            Assert.Equal(Effect.OutlineShadow, p.Effect);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/emboss"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/shadow"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outline"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outlineShadow"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/imprint"));

            p.Effect = Effect.Engrave;
            Assert.Equal(Effect.Engrave, p.Effect);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/emboss"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/shadow"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outline"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/outlineShadow"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/imprint"));
        }

        [Fact]
        public void ExpansionAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Null(p.ExpansionScale);

            p.ExpansionScale = 200;
            Assert.Equal(200, p.ExpansionScale);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/w"));

            p.ExpansionScale = null;
            Assert.Null(p.ExpansionScale);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/w"));

            Assert.Throws<ArgumentOutOfRangeException>(() => p.ExpansionScale = -1);
        }

        [Fact]
        public void PositionAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Null(p.Position);

            p.Position = 50;
            Assert.Equal(50, p.Position);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/position"));
            Assert.Equal("100", p.Xml.RemoveNamespaces().XPathSelectElement("//rPr/position").Attribute("val")?.Value);

            p.Position = null;
            Assert.Null(p.Position);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/position"));
        }

        [Fact]
        public void SuperscriptAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.False(p.Superscript);
            Assert.False(p.Subscript);

            p.Superscript = true;
            Assert.True(p.Superscript);
            Assert.False(p.Subscript);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vertAlign"));

            p.Superscript = false;
            Assert.False(p.Superscript);
            Assert.False(p.Subscript);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vertAlign"));
        }

        [Fact]
        public void SuperscriptAffectSubscript()
        {
            // Default
            var p = new Paragraph();
            Assert.False(p.Superscript);
            Assert.False(p.Subscript);

            p.Superscript = true;
            Assert.True(p.Superscript);
            Assert.False(p.Subscript);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vertAlign"));

            p.Subscript = true;
            Assert.False(p.Superscript);
            Assert.True(p.Subscript);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vertAlign"));

            p.Superscript = false;
            Assert.False(p.Superscript);
            Assert.True(p.Subscript);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vertAlign"));

            p.Subscript = false;
            Assert.False(p.Superscript);
            Assert.False(p.Subscript);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/vertAlign"));
        }

        [Fact]
        public void LineSpacingAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Null(p.LineSpacing);

            p.LineSpacing = 12.5;
            Assert.Equal(12.5, p.LineSpacing);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));

            p.LineSpacing = null;
            Assert.Null(p.LineSpacing);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));
        }

        [Fact]
        public void LineSpacingBeforeAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Null(p.LineSpacingBefore);

            p.LineSpacingBefore = 12.5;
            Assert.Equal(12.5, p.LineSpacingBefore);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));

            p.LineSpacingBefore = null;
            Assert.Null(p.LineSpacingBefore);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));
        }

        [Fact]
        public void LineSpacingAfterAddsRemovesElement()
        {
            // Default
            var p = new Paragraph();
            Assert.Null(p.LineSpacingAfter);

            p.LineSpacingAfter = 12.5;
            Assert.Equal(12.5, p.LineSpacingAfter);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));

            p.LineSpacingAfter = null;
            Assert.Null(p.LineSpacingAfter);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));
        }

        [Fact]
        public void LineSpacingRemovesWhenNoAttributes()
        {
            var p = new Paragraph {LineSpacing = 10, LineSpacingBefore = 12, LineSpacingAfter = 15};

            Assert.Equal(10, p.LineSpacing);
            Assert.Equal(12, p.LineSpacingBefore);
            Assert.Equal(15, p.LineSpacingAfter);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));

            p.LineSpacingAfter = null;
            Assert.Null(p.LineSpacingAfter);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));

            p.LineSpacingBefore = null;
            Assert.Null(p.LineSpacingBefore);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));

            p.LineSpacing = null;
            Assert.Null(p.LineSpacing);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/spacing"));
        }

        [Fact]
        public void StrikeThroughAddsRemovesElement()
        {
            var p = new Paragraph();
            Assert.Equal(Strikethrough.None, p.StrikeThrough);

            p.StrikeThrough = Strikethrough.Strike;
            Assert.Equal(Strikethrough.Strike, p.StrikeThrough);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/strike"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/dstrike"));

            p.StrikeThrough = Strikethrough.DoubleStrike;
            Assert.Equal(Strikethrough.DoubleStrike, p.StrikeThrough);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/strike"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/dstrike"));

            p.StrikeThrough = Strikethrough.None;
            Assert.Equal(Strikethrough.None, p.StrikeThrough);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/strike"));
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/dstrike"));
        }

        [Fact]
        public void StyleAddsRemovesElement()
        {
            var p = new Paragraph();
            Assert.Equal("Normal", p.StyleName);

            p.StyleName = "Body";
            Assert.Equal("Body", p.StyleName);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/pStyle"));

            p.StyleName = null;
            Assert.Equal("Normal", p.StyleName);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/pStyle"));

            p.StyleName = "";
            Assert.Equal("Normal", p.StyleName);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//pPr/pStyle"));
        }

        [Fact]
        public void UnderlineStyleAddsRemovesElement()
        {
            var p = new Paragraph();
            Assert.Equal(UnderlineStyle.None, p.UnderlineStyle);

            p.UnderlineStyle = UnderlineStyle.SingleLine;
            Assert.Equal(UnderlineStyle.SingleLine, p.UnderlineStyle);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/u"));

            p.UnderlineStyle = UnderlineStyle.DoubleLine;
            Assert.Equal(UnderlineStyle.DoubleLine, p.UnderlineStyle);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/u"));

            p.UnderlineStyle = UnderlineStyle.None;
            Assert.Equal(UnderlineStyle.None, p.UnderlineStyle);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("//rPr/u"));
        }

        [Fact]
        public void SetTextReplacesContents()
        {
            var p = new Paragraph();
            p.InsertText("This is a test. ");
            p.AppendLine("Will it work?").Bold();
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

            Paragraph p = new Paragraph("Test").Bold();
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
                .Append(" is ").Bold()
                .Append("a ").Italic()
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

            p.GetFormatting(true);
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

            p.GetFormatting(true);
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

            p.Append("This is a test").Bold();
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//r"));

            var lastRun = p.Xml.Elements(Name.Run).Last();
            var properties = lastRun.Element(Name.RunProperties);
            Assert.Single(properties.Elements(Name.Bold));
        }

        [Fact]
        public void FormattingChangesSkipLineBreaks()
        {
            var p = new Paragraph();

            p.AppendLine("This is a test").Bold();
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
