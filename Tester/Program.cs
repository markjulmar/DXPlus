using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DXPlus;

namespace Tester
{
    public static class Program
    {
        public static void Main()
        {
            string fn = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.docx");
            using var doc = Document.Create(fn);

            WriteTitle(doc);

            AddVideoToDoc(doc);
            AddImageToDoc(doc);
            AddHeaderAndFooter(doc);
            CreateTableWithList(doc);

            doc.Save();
            Console.WriteLine("Wrote document");
        }

        private static void AddVideoToDoc(IDocument doc)
        {
            doc.AddRange(new[] {
                new Paragraph("This is a video."),
                new Paragraph().WithProperties(new() {Alignment = Alignment.Center})
                    .Append(doc.CreateVideo(
                        "video-placeholder.png",
                        new Uri("https://www.microsoft.com/en-us/videoplayer/embed/RWwMdr", UriKind.Absolute),
                        400, 225)),
                new Paragraph("And a closing paragraph.")
            });
        }

        private static void WriteTitle(IDocument doc)
        {
            doc.AddRange(new[] {"Introduction", "This is some text"});
            doc.Paragraphs.First()
             .Style(HeadingType.Heading1)
             .InsertBefore(new Paragraph("This is a title")
                 .Style(HeadingType.Title))
             .Append(new Paragraph($"Last edited at {DateTime.Now.ToShortDateString()} by M. Smith")
                 .Style(HeadingType.Subtitle));
        }

        private static void AddImageToDoc(IDocument doc)
        {
            doc.AddParagraph();

            var img = doc.CreateImage(@"test.svg");
            var p = doc.Add("This is a picture:");
            p.Append(img.CreatePicture(string.Empty, string.Empty));

            var im2 = doc.CreateImage(@"test2.png");
            p.Append(im2.CreatePicture(string.Empty, string.Empty));

            // Add with different size.
            p = doc.Add("And a final pic (dup of svg!):");
            p.Append(img.CreatePicture(50, 50));
        }

        private static void AddHeaderAndFooter(IDocument doc)
        {
            var mainSection = doc.Sections.First();
            var header = mainSection.Headers.Default;

            var p1 = header.MainParagraph;
            p1.SetText("This is some text - ");
            p1.AddPageNumber(PageNumberFormat.Normal);
        }

        static void CreateTableWithList(IDocument doc)
        {
            doc.AddPageBreak();
            doc.AddParagraph("This is a table.");

            var table = new Table(rows: 2, columns: 2) {Design = TableDesign.None};
            table.SetOutsideBorders(new Border(BorderStyle.Single, Uom.FromPoints(1))); //1pt
            table.SetInsideBorders(new Border(BorderStyle.Single, Uom.FromPoints(1.5))); // 1.5pt

            doc.Add(table);

            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);
            var rows = table.Rows.ToList();

            foreach (var row in rows)
            {
                for (int col = 0; col < row.ColumnCount; col++)
                {
                    AddList(row.Cells[col].Paragraphs.First(), 
                        nd, Enumerable.Range(1,5).Select(n => $"Item {n}"));    
                    //row.Cells[col].Text = "Hello";
                }
            }

            table.AutoFit();
        }

        private static void AddList(Paragraph paragraph, NumberingDefinition nd, IEnumerable<string> terms)
        {
            int index = 0;
            foreach (var text in terms)
            {
                if (index++ == 0)
                    paragraph.ListStyle(nd, 0);
                else
                    paragraph = paragraph.AddParagraph().ListStyle(nd,0);
                paragraph.SetText(text);
            }
        }
    }
}
