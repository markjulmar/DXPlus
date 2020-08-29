using System.Xml.XPath;
using Xunit;

namespace DXPlus.Tests
{
    public class ParagraphPropertyTests
    {
        [Fact]
        public void KeepWithNextAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.False(p.KeepWithNext);

            p.KeepWithNext = true;
            Assert.True(p.KeepWithNext);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("keepNext"));

            p.KeepWithNext = true;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("keepNext"));

            p.KeepWithNext = false;
            Assert.False(p.KeepWithNext);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("keepNext"));
        }

        [Fact]
        public void KeepWLinesTogetherAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.False(p.KeepLinesTogether);

            p.KeepLinesTogether = true;
            Assert.True(p.KeepLinesTogether);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("keepLines"));

            p.KeepLinesTogether = true;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("keepLines"));

            p.KeepLinesTogether = false;
            Assert.False(p.KeepLinesTogether);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("keepLines"));
        }

        [Fact]
        public void AlignmentGetAndSetAreAligned()
        {
            var p = new ParagraphProperties();
            Assert.Equal(Alignment.Left, p.Alignment);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("jc"));

            p.Alignment = Alignment.Center;
            Assert.Equal(Alignment.Center, p.Alignment);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("jc"));

            p.Alignment = Alignment.Center;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("jc"));

            p.Alignment = Alignment.Left;
            Assert.Equal(Alignment.Left, p.Alignment);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("jc"));
        }

        [Fact]
        public void StyleGetAndSetAreAligned()
        {
            var p = new ParagraphProperties();
            Assert.Equal("Normal", p.StyleName);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("pStyle"));

            p.StyleName = "Test";
            Assert.Equal("Test", p.StyleName);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("pStyle"));

            p.StyleName = "Test";
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("pStyle"));

            p.StyleName = "Normal";
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("pStyle"));
        }

        [Fact]
        public void LineSpacingAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.Null(p.LineSpacing);

            p.LineSpacing = 12.5;
            Assert.Equal(12.5, p.LineSpacing);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing[@line='250']"));

            p.LineSpacing = null;
            Assert.Null(p.LineSpacing);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));
        }

        [Fact]
        public void LineSpacingBeforeAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.Null(p.LineSpacingBefore);

            p.LineSpacingBefore = 12.5;
            Assert.Equal(12.5, p.LineSpacingBefore);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing[@before='250']"));

            p.LineSpacingBefore = null;
            Assert.Null(p.LineSpacingBefore);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));
        }

        [Fact]
        public void LineSpacingAfterAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.Null(p.LineSpacingAfter);

            p.LineSpacingAfter = 16.75;
            Assert.Equal(16.75, p.LineSpacingAfter);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing[@after='335']"));

            p.LineSpacingAfter = null;
            Assert.Null(p.LineSpacingAfter);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));
        }

        [Fact]
        public void LineSpacingRemovesWhenNoAttributes()
        {
            var p = new ParagraphProperties { LineSpacing = 10, LineSpacingBefore = 12, LineSpacingAfter = 15 };

            Assert.Equal(10, p.LineSpacing);
            Assert.Equal(12, p.LineSpacingBefore);
            Assert.Equal(15, p.LineSpacingAfter);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));

            p.LineSpacingAfter = null;
            Assert.Null(p.LineSpacingAfter);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));

            p.LineSpacingBefore = null;
            Assert.Null(p.LineSpacingBefore);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));

            p.LineSpacing = null;
            Assert.Null(p.LineSpacing);
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("spacing"));
        }

        [Fact]
        public void IndentLeftAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.Equal(0, p.IndentationLeft);

            p.IndentationLeft = 15.225;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind[@left='304.5']"));
            Assert.Equal(15.22, p.IndentationLeft);

            p.IndentationLeft = 0;
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Equal(0, p.IndentationLeft);
        }

        [Fact]
        public void IndentRightAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.Equal(0, p.IndentationRight);

            p.IndentationRight = 15.225;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind[@right='304.5']"));
            Assert.Equal(15.22, p.IndentationRight);

            p.IndentationRight = 0;
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Equal(0, p.IndentationRight);
        }

        [Fact]
        public void IndentFirstLineAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.Equal(0, p.IndentationFirstLine);

            p.IndentationFirstLine = 15.225;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind[@firstLine='304.5']"));
            Assert.Equal(15.22, p.IndentationFirstLine);

            p.IndentationFirstLine = 0;
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Equal(0, p.IndentationFirstLine);
        }

        [Fact]
        public void IndentHangingAddsRemovesElement()
        {
            var p = new ParagraphProperties();
            Assert.Equal(0, p.IndentationHanging);

            p.IndentationHanging = 15.225;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind[@hanging='304.5']"));
            Assert.Equal(15.22, p.IndentationHanging);

            p.IndentationHanging = 0;
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Equal(0, p.IndentationHanging);
        }

        [Fact]
        public void IndentHangingAffectsFirstLine()
        {
            var p = new ParagraphProperties();

            p.IndentationHanging = 10;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));

            p.IndentationFirstLine = 12;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Equal(0, p.IndentationHanging);

            p.IndentationHanging = 15;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Equal(0, p.IndentationFirstLine);
        }

        [Fact]
        public void IndentsShareElement()
        {
            var p = new ParagraphProperties();

            p.IndentationHanging = 10;
            p.IndentationLeft = 15;
            p.IndentationRight = 20;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            Assert.Equal(10, p.IndentationHanging);
            Assert.Equal(15, p.IndentationLeft);
            Assert.Equal(20, p.IndentationRight);
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind[@left='300' and @right='400' and @hanging='200']"));

            p.IndentationHanging = 0;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            p.IndentationLeft = 0;
            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
            p.IndentationRight = 0;
            Assert.Empty(p.Xml.RemoveNamespaces().XPathSelectElements("ind"));
        }

    }
}
