using System.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class SectionPropertyTests
    {
        [Fact]
        public void CanReadNewDocProperties()
        {
            var doc = Document.Create();
            var properties = doc.Sections.Single().Properties;

            Assert.Equal(PageSize.LetterHeight, properties.PageHeight);
            Assert.Equal(PageSize.LetterWidth, properties.PageWidth);
            Assert.Equal(Orientation.Portrait, properties.Orientation);

            Assert.Equal(1440, properties.BottomMargin);
            Assert.Equal(1440, properties.TopMargin);
            Assert.Equal(1440, properties.LeftMargin);
            Assert.Equal(1440, properties.RightMargin);
        }
    }
}
