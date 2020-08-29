using System.Xml.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class TextBlockTests
    {
        [Fact]
        public void CtorWithTextHasProperValue()
        {
            string text = "This is a test";
            var xml = new XElement(Name.Text, text);

            var tb = new TextBlock(xml, 0);
            Assert.Equal(text, tb.Value);
        }

        [Fact]
        public void CtorWithTabHasProperValue()
        {
            var xml = new XElement(Namespace.Main + "tab");
            var tb = new TextBlock(xml, 0);
            Assert.Equal("\t", tb.Value);
        }

        [Fact]
        public void CtorWithBreakHasProperValue()
        {
            var xml = new XElement(Namespace.Main + "br");
            var tb = new TextBlock(xml, 0);
            Assert.Equal("\n", tb.Value);
        }

        [Fact]
        public void CtorWithCarriageReturnHasProperValue()
        {
            var xml = new XElement(Namespace.Main + "cr");
            var tb = new TextBlock(xml, 0);
            Assert.Equal("\n", tb.Value);
        }

        [Fact]
        public void CtorWithDelTextHasProperValueAndLength()
        {
            string text = "Some deleted text.";
            var xml = new XElement(Namespace.Main + "delText", text);
            var tb = new TextBlock(xml, 0);
            Assert.Equal(text, tb.Value);
        }

        [Fact]
        public void SplitTextReturnsBothSides()
        {
            string text = "This is some text";
            var tb = new TextBlock(new XElement(Name.Text, text), 10);
            Assert.Equal(10, tb.StartIndex);
            Assert.Equal(text.Length+10, tb.EndIndex);

            var results = tb.Split(15);
            Assert.Equal("This ", results[0].Value);
            Assert.Equal("is some text", results[1].Value);
        }
    }
}
