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
        public void ExtensionReturnsAllListItems()
        {
            using var doc = Document.Create();
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            doc.AddParagraph("Separator paragraph.");
            doc.AddParagraph("Separator paragraph.");
            doc.AddParagraph("Separator paragraph.");

            int i;
            for (i = 0; i < 3; i++)
            {
                doc.AddParagraph($"Item #{i+1}").ListStyle(nd);
            }

            doc.AddParagraph("Separator paragraph.");
            doc.AddParagraph("Separator paragraph.");
            doc.AddParagraph("Separator paragraph.");

            for (; i < 5; i++)
            {
                doc.AddParagraph($"Item #{i + 1}").ListStyle(nd);
            }

            doc.AddParagraph("Separator paragraph.");
            doc.AddParagraph("Separator paragraph.");
            doc.AddParagraph("Separator paragraph.");

            var list = doc.GetListById(nd.Id).ToList();
            Assert.Equal(5, list.Count);
            for (i = 0; i < list.Count; i++)
            {
                Assert.Equal($"Item #{i + 1}", list[i].Text);
            }
        }

        [Fact]
        public void CanFindListItemIndex()
        {
            using var doc = Document.Create();
            var nd = doc.NumberingStyles.Create(NumberingFormat.Numbered);

            doc.AddParagraph("Separator paragraph.");

            for (int i = 0; i < 5; i++)
            {
                doc.AddParagraph($"Item #{i + 1}").ListStyle(nd);
            }

            doc.AddParagraph("Separator paragraph.");

            var p = doc.Paragraphs.Single(p => p.Text == "Item #2");
            Assert.Equal(1, p.GetListIndex());

            p = doc.Paragraphs.Single(p => p.Text == "Item #5");
            Assert.Equal(4, p.GetListIndex());
        }

        [Fact]
        public void ListAddsToHeaderPart()
        {
            using var doc = Document.Create();
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            var section = doc.Sections.Single();
            var header = section.Headers.Default;

            var p = header.AddParagraph("List in paragraph").ListStyle(nd);
            Assert.NotNull(p.PackagePart);
            Assert.Equal(p.PackagePart, header.PackagePart);
        }
    }
}
