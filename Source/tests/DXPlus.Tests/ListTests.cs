using System.Linq;
using Newtonsoft.Json.Serialization;
using Xunit;

namespace DXPlus.Tests
{
    public class ListTests
    {
        [Fact]
        public void ListHasEnumerableItems()
        {
            var list = new List(NumberingFormat.Bulleted);
            list.AddItem("1")
                .AddItem("2")
                .AddItem("3");

            Assert.Equal(3, list.Items.Count);
            Assert.Equal("2", list.Items[1].Paragraph.Text);
        }

        [Fact]
        public void PackagePartSetWhenAddedToDoc()
        {
            List list = new List(NumberingFormat.Bulleted);
            Assert.Null(list.PackagePart);

            var doc = Document.Create();
            doc.AddList(list);
            Assert.NotNull(list.PackagePart);
            Assert.Empty(doc.Lists);
        }

        [Fact]
        public void ListAddsElementsToDoc()
        {
            var doc = Document.Create();
            doc.AddList(new List(NumberingFormat.Bulleted, new[] { "Item 1", "Item 2", "Item 3" }));
            Assert.Single(doc.Lists);

            doc.AddList(new List(NumberingFormat.Numbered, new [] { "One" }));
            Assert.Equal(2, doc.Lists.Count());
        }

        [Fact]
        public void ListAddsElementsToDocAfterInsert()
        {
            List list = new List(NumberingFormat.Bulleted);
            var doc = Document.Create();
            var l2 = doc.AddList(list);
            Assert.Empty(doc.Lists);

            l2.AddItem("2");
            Assert.Single(doc.Lists);
        }

        [Fact]
        public void ListAddsToHeaderPart()
        {
            List list = new List(NumberingFormat.Bulleted, new [] { "Test"} );
            Assert.Null(list.PackagePart);

            var doc = Document.Create();
            var section = doc.Sections.Single();

            section.Headers.Default.Add();
            var header = section.Headers.Default;

            header.AddList(list);
            Assert.NotNull(list.PackagePart);
            Assert.Single(header.Lists);
        }
    }
}
