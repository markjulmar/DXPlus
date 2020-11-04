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

            Assert.Equal(792, properties.PageHeight);
            Assert.Equal(612, properties.PageWidth);
            Assert.Equal(Orientation.Portrait, properties.Orientation);

            Assert.Equal(72, properties.BottomMargin);
            Assert.Equal(72, properties.BottomMargin);
            Assert.Equal(72, properties.BottomMargin);
            Assert.Equal(72, properties.BottomMargin);
        }
    }
}
