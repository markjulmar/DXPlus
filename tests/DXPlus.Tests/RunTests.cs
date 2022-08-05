using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Internal;
using Xunit;

namespace DXPlus.Tests
{
    public class RunTests
    {
        [Fact]
        public void PublicCtorCreatesTextRun()
        {
            string text = "This is a test";
            var run = new Run(text, new Formatting());
            Assert.Equal(text, run.Text);

            Assert.Equal(Name.Run, run.Xml.Name);
            Assert.Single(run.Xml.Elements(Name.Text));
            Assert.Single(run.Xml.Elements(Name.RunProperties));
        }

        [Fact]
        public void SplitMiddleReturnsBothSides()
        {
            string text = "This is a test.";
            var r = new Run(null, null, new XElement(Name.Run, new XElement(Name.Text, text)));
            Assert.Equal(text, r.Text);

            var (leftElement, rightElement) = r.Split(5);

            Assert.Equal("This ", leftElement.Value);
            Assert.Equal("is a test.", rightElement.Value);
        }

        [Fact]
        public void SplitAtZeroReturnsRightSide()
        {
            string text = "This is a test.";
            var r = new Run(null, null, new XElement(Name.Run, new XElement(Name.Text, text)));
            Assert.Equal(text, r.Text);

            var (leftElement, rightElement) = r.Split(0);

            Assert.Null(leftElement);
            Assert.Equal("This is a test.", rightElement.Value);
        }

        [Fact]
        public void SplitAtLengthReturnsLeftSide()
        {
            string text = "This is a test.";
            var r = new Run(null, null, new XElement(Name.Run, new XElement(Name.Text, text)));
            Assert.Equal(text, r.Text);

            var (leftElement, rightElement) = r.Split(text.Length);

            Assert.Null(rightElement);
            Assert.Equal("This is a test.", leftElement.Value);
        }

        [Fact]
        public void FirstRunSplitInsertAtLengthReturnsLeftSide()
        {
            string text = "Test";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(null, null, e);
            Assert.Equal(text, r.Text);
            var (leftElement, rightElement) = r.Split(text.Length);

            Assert.Null(rightElement);
            Assert.Equal("Test", leftElement.Value);
        }

        [Fact]
        public void TextSkipDeletedText()
        {
            string xml =
                @"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
					    <w:delText xml:space=""preserve"">deleted </w:delText>
				    </w:r>";
            Run run = new Run(null, null, XElement.Parse(xml));
            Assert.Equal("", run.Text);

            Assert.Single(run.Elements);
            Assert.Equal(run.Elements.First().ElementType, RunTextType.DeletedText);
            Assert.Equal("deleted ", ((DeletedText)run.Elements.First()).Value);
        }

        [Fact]
        public void MergeReplacesWhenNoFormattingPresent()
        {
            var r1 = new Run("This is a test");
            Assert.Null(r1.Properties);

            var f = new Formatting {Bold = true, Color = Color.Red};
            r1.MergeFormatting(f);
            Assert.Equal(f, r1.Properties);
        }

        [Fact]
        public void MergeAddsToFormatting()
        {
            var r1 = new Run("This is a test", new Formatting { Italic = true, Font = new FontValue("Arial"), FontSize = 16 });
            Assert.NotNull(r1.Properties);

            var f = new Formatting { Bold = true, Color = Color.Red };
            var f2 = new Formatting
                {Bold = true, Color = Color.Red, Italic = true, Font = new FontValue("Arial"), FontSize = 16};
            r1.MergeFormatting(f);
            Assert.Equal(f2, r1.Properties);
        }

        [Fact]
        public void CanGetParentFromTableRun()
        {
            var doc = Document.Create();

            doc.Add("First paragraph");
            doc.Add(new Table(new[,] {{"H1", "H2"}, {"C1", "C2"}, {"C3", "C4"}}));

            Assert.Equal(2, doc.Blocks.Count());
            Assert.Equal(7, doc.Paragraphs.Count());

            var p = doc.Paragraphs.Skip(3).First();
            Assert.Equal("C1", p.Text);

            var r = p.Runs.First();

            Assert.NotNull(r.Parent);
            Assert.Equal(p, r.Parent);

        }

        [Fact]
        public void CombinationOperatorMergesFormatting()
        {
            var f1 = new Formatting() {Bold = true};
            var f2 = new Formatting() {Italic = true};
            var f3 = f1 + f2;
            Assert.Equal(new Formatting() { Bold = true, Italic = true }, f3);
        }

        [Fact]
        public void SubtractionOperatorRemovesFormatting()
        {
            var f1 = new Formatting() { Bold = true, Italic = true, Color = Color.Red };
            var f2 = new Formatting() { Italic = true, Effect = Effect.Emboss, Color = Color.Blue };
            var f3 = f1 - f2;
            Assert.Equal(new Formatting() { Bold = true, Color = Color.Red }, f3);
        }
    }
}
