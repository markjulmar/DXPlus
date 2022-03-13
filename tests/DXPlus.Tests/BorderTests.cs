using System.Drawing;
using System.Xml.XPath;
using Xunit;

namespace DXPlus.Tests
{
    public class BorderTests
    {
        [Fact]
        public void RunPropertiesUsesBdrTag()
        {
            var props = new Formatting();
            Assert.Null(props.Border);

            props.Border = new Border(BorderStyle.Single, Uom.FromPoints(1));
            Assert.NotNull(props.Border);

            Assert.Single(props.Xml.RemoveNamespaces().XPathSelectElements("bdr"));
        }

        [Fact]
        public void TablePropertiesUsesTblBdrTag()
        {
            var t = new Table();
            Assert.Null(t.LeftBorder);

            t.LeftBorder = new Border(BorderStyle.Single, Uom.FromPoints(1));
            Assert.NotNull(t.LeftBorder);

            Assert.Single(t.Xml.RemoveNamespaces().XPathSelectElements("//tblBorders/left"));
        }

        [Fact]
        public void ParagraphPropertiesUsespBdrTag()
        {
            var p = new Paragraph();
            Assert.Null(p.Properties.RightBorder);

            p.Properties.RightBorder = new Border(BorderStyle.Single, Uom.FromPoints(1));
            Assert.NotNull(p.Properties.RightBorder);

            Assert.Single(p.Xml.RemoveNamespaces().XPathSelectElements("//pBdr/right"));
        }

        [Fact]
        public void SetSameObjectHasNoEffect()
        {
            var props = new ParagraphProperties();
            var border = new Border(BorderStyle.Single, Uom.FromPoints(1));

            props.TopBorder = border;
            border = props.TopBorder;
            Assert.Equal(border, props.TopBorder);
            props.TopBorder = border;
            Assert.Equal(border, props.TopBorder);
        }

        [Fact]
        public void AssignSameObjectToMultiplesCopies()
        {
            var props = new ParagraphProperties();
            var border = new Border(BorderStyle.Single, Uom.FromPoints(1.5));

            props.TopBorder = border;
            props.BottomBorder = border;
            props.LeftBorder = border;
            props.RightBorder = border;

            Assert.Equal(12, props.TopBorder.Size);
            Assert.Equal(12, props.BottomBorder.Size);
            Assert.Equal(12, props.LeftBorder.Size);
            Assert.Equal(12, props.RightBorder.Size);
            Assert.NotEqual(props.TopBorder, border);
        }

        [Fact]
        public void NoValuesRemovesBorder()
        {
            var props = new ParagraphProperties();
            Assert.Null(props.TopBorder);

            var border = new Border(BorderStyle.None, 5) {Size = null};
            props.TopBorder = border;
            Assert.Null(props.TopBorder);
        }

        [Fact]
        public void ChangeValuesRemovesBorder()
        {
            var props = new ParagraphProperties() {TopBorder = new Border(BorderStyle.None, 5)};
            Assert.NotNull(props.TopBorder);

            var topBorder = props.TopBorder;
            topBorder.Color = new ColorValue(ThemeColor.Accent1, null, null);
            Assert.NotNull(props.TopBorder);

            props.TopBorder.Size = null;
            Assert.NotNull(props.TopBorder);

            topBorder.Color = new ColorValue(Color.Empty);
            Assert.Null(props.TopBorder);
        }
    }
}
