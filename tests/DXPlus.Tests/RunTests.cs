using System.Xml;
using System.Xml.Linq;
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
        public void SplitInsertRunReturnsBothSides()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(null, null, e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitAtIndex(10);

            Assert.Equal("This ", results[0].Value);
            Assert.Equal("is a test.", results[1].Value);
        }

        [Fact]
        public void SplitDeleteRunReturnsBothSides()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(null, null, e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitAtIndex(10);

            Assert.Equal("This ", results[0].Value);
            Assert.Equal("is a test.", results[1].Value);
        }

        [Fact]
        public void SplitInsertAtZeroRunReturnsRightSide()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(null, null, e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitAtIndex(5);

            Assert.Null(results[0]);
            Assert.Equal("This is a test.", results[1].Value);
        }

        [Fact]
        public void SplitDeleteAtZeroRunReturnsRightSide()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(null, null, e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitAtIndex(5);

            Assert.Null(results[0]);
            Assert.Equal("This is a test.", results[1].Value);
        }

        [Fact]
        public void SplitInsertAtLengthRunReturnsLeftSide()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(null, null, e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitAtIndex(5 + text.Length);

            Assert.Null(results[1]);
            Assert.Equal("This is a test.", results[0].Value);
        }

        [Fact]
        public void SpliDeleteAtLengthRunReturnsLeftSide()
        {
            string text = "This is a test.";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(null, null, e, 5);
            Assert.Equal(text, r.Text);
            var results = r.SplitAtIndex(5 + text.Length);

            Assert.Null(results[1]);
            Assert.Equal("This is a test.", results[0].Value);
        }

        [Fact]
        public void FirstRunSplitInsertAtLengthReturnsLeftSide()
        {
            string text = "Test";
            var e = new XElement(Name.Run, new XElement(Name.Text, text));

            Run r = new Run(null, null, e, 0);
            Assert.Equal(text, r.Text);
            var results = r.SplitAtIndex(text.Length);

            Assert.Null(results[1]);
            Assert.Equal("Test", results[0].Value);
        }

    }
}
