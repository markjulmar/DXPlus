﻿using System.Linq;
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
            Assert.Equal(1, p.GetListNumberingDefinitionId());
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
        public void CanGetNumberingDefinitionFromParagraph()
        {
            using var doc = Document.Create();
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            doc.AddParagraph("Separator paragraph.");

            for (int i = 0; i < 3; i++)
            {
                doc.AddParagraph($"Item #{i + 1}").ListStyle(nd);
            }

            var paragraphs = doc.Paragraphs.ToList();
            Assert.Null(paragraphs[0].GetListNumberingDefinition());
            Assert.Equal(nd.Id, paragraphs[1].GetListNumberingDefinitionId());
            Assert.Equal(nd.Id, paragraphs[2].GetListNumberingDefinition().Id);
            Assert.Equal(nd.Id, paragraphs[3].GetListNumberingDefinition().Id);
        }

        [Fact]
        public void CanGetNumberingDefinitionFromParagraphWhenNoIdIsPresent()
        {
            using var doc = Document.Create();
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            doc.AddParagraph("Separator paragraph.");

            doc.AddParagraph($"Item #1").ListStyle(nd);
            doc.AddParagraph($"Item #2").ListStyle();
            doc.AddParagraph($"Item #3").ListStyle();

            var paragraphs = doc.Paragraphs.ToList();
            Assert.Null(paragraphs[0].GetListNumberingDefinition());
            Assert.Equal(nd.Id, paragraphs[1].GetListNumberingDefinitionId());
            Assert.Equal(nd.Id, paragraphs[2].GetListNumberingDefinition().Id);
            Assert.Equal(nd.Id, paragraphs[3].GetListNumberingDefinition().Id);
        }

        [Fact]
        public void CanGetNumberingDefinitionFromParagraphWhenMultipleAreUsed()
        {
            using var doc = Document.Create();
            var n1 = doc.NumberingStyles.Create(NumberingFormat.Bullet);
            var n2 = doc.NumberingStyles.Create(NumberingFormat.Numbered);

            doc.AddParagraph("Separator paragraph.");

            doc.AddParagraph($"Item #1").ListStyle(n1);
            doc.AddParagraph($"Item #2").ListStyle();
            doc.AddParagraph($"Item #1").ListStyle(n2);
            doc.AddParagraph($"Item #2").ListStyle();

            doc.AddParagraph("Separator paragraph.");

            var paragraphs = doc.Paragraphs.ToList();

            Assert.Null(paragraphs[0].GetListNumberingDefinition());
            Assert.Null(paragraphs[5].GetListNumberingDefinition());

            Assert.Equal(n1.Id, paragraphs[1].GetListNumberingDefinitionId());
            Assert.Equal(n1.Id, paragraphs[2].GetListNumberingDefinition().Id);
            Assert.Equal(n2.Id, paragraphs[3].GetListNumberingDefinition().Id);
            Assert.Equal(n2.Id, paragraphs[4].GetListNumberingDefinitionId());
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
