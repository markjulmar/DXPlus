using System.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class ListTests
    {
        [Fact]
        public void ListStyleAddsNumberingDefinitionId()
        {
            using var doc = Document.Create();
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            var p = doc.AddParagraph("1").ListStyle(nd);

            Assert.True(p.IsListItem());
            Assert.Equal(1, p.GetListNumId());
            Assert.Equal(0, p.GetListLevel());
        }

        [Fact]
        public void ChildItemsHaveListStyle()
        {
            using var doc = Document.Create();
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            var p = doc.AddParagraph("Starting list").ListStyle(nd);
            p = p.AddParagraph("With another paragraph").ListStyle();

            Assert.Equal("ListParagraph", p.Properties.StyleName);
        }

        [Fact]
        public void ListAddsToHeaderPart()
        {
            using var doc = Document.Create();
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            var section = doc.Sections.Single();

            section.Headers.Default.Add();
            var header = section.Headers.Default;

            var p = header.AddParagraph("List in paragraph").ListStyle(nd);
            Assert.NotNull(p.PackagePart);
            Assert.Equal(p.PackagePart, header.PackagePart);
        }
    }
}
