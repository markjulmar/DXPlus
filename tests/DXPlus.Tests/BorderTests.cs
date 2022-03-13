using Xunit;

namespace DXPlus.Tests
{
    public class BorderTests
    {
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
    }
}
