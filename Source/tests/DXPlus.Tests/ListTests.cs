using System.Linq;
using Xunit;

namespace DXPlus.Tests
{
    public class ListTests
    {
        private const string Filename = "test.docx";

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
            var l2 = doc.AddList(list);
            Assert.NotNull(l2.PackagePart);

            var l3 = doc.InsertList(1, list);
            Assert.NotNull(l3.PackagePart);
            Assert.Empty(doc.Lists);
        }

        [Fact]
        public void ListAddsElementsToDoc()
        {
            List list = new List(NumberingFormat.Bulleted);
            list.AddItem("Item 1");

            var doc = Document.Create();
            var l2 = doc.AddList(list);
            Assert.Single(doc.Lists);

            var l3 = doc.AddList(list);
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
            List list = new List(NumberingFormat.Bulleted)
                            .AddItem("Test");
            Assert.Null(list.PackagePart);

            var doc = Document.Create();

            doc.Headers.Default.Add();
            var header = doc.Headers.Default;

            var l2 = header.AddList(list);
            Assert.NotNull(l2.PackagePart);

            var l3 = doc.AddList(list);

            Assert.Single(doc.Lists);
            Assert.Single(header.Lists);
            Assert.NotEqual(l2.PackagePart, l3.PackagePart);
        }
    }
}
