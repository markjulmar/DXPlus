using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace DXPlus.Tests
{
    public class ListTests
    {
        private const string Filename = "test.docx";

        [Fact]
        public void ListHasEnumerableItems()
        {
            var list = new List(ListItemType.Bulleted);
            list.AddItem("1")
                .AddItem("2")
                .AddItem("3");

            Assert.Equal(3, list.Items.Count);
            Assert.Equal("2", list.Items[1].Paragraph.Text);
        }

    }
}
