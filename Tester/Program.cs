using System;
using System.Collections.Generic;
using System.Drawing;
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
                "This is a video.",
                new Paragraph().WithProperties(new() {Alignment = Alignment.Center})
                    .Add(doc.CreateVideo(
                        "video-placeholder.png",
                        new Uri("https://www.youtube.com/watch?v=5-gF-tmblA8", UriKind.Absolute),
                        400, 225)),
                new Paragraph(new [] {
                            "with a ", 
                            new Run("boxed", 
                                new Formatting { Border = new Border(BorderStyle.Dotted, Uom.FromPoints(1))}),
                            " caption.",
                        }),
                "And a closing paragraph.",

                new Paragraph("One more time with a border")
                    .SetOutsideBorders(new Border(BorderStyle.DoubleWave, 5))
            });
        }

        private static void WriteTitle(IDocument doc)
        {
            doc.AddRange(new[] {"Introduction", "This is some text"});
            doc.Paragraphs.First()
             .Style(HeadingType.Heading1)
             .InsertBefore(new Paragraph("This is a title").Style(HeadingType.Title))
             .InsertAfter(new Paragraph($"Last edited at {DateTime.Now.ToShortDateString()} by M. Smith").Style(HeadingType.Subtitle));
        }

        private static void AddImageToDoc(IDocument doc)
        {
            doc.AddParagraph();

            var img = doc.CreateImage(@"test.svg");
            var p = doc.Add("This is a picture:");
            p.Add(img.CreatePicture(string.Empty, string.Empty));

            var im2 = doc.CreateImage(@"test2.png");
            p.Add(im2.CreatePicture(string.Empty, string.Empty));

            // Add with different size.
            p = doc.Add("And a final pic (dup of svg!):");
            p.Add(img.CreatePicture(50, 50));
        }

        private static void AddHeaderAndFooter(IDocument doc)
        {
            var mainSection = doc.Sections.First();
            var header = mainSection.Headers.Default;

            var p1 = header.MainParagraph;
            p1.Text = "This is some text - ";
            p1.AddPageNumber(PageNumberFormat.Normal);
        }

        static void CreateTableWithList(IDocument doc)
        {
            doc.AddPageBreak();
            doc.Add("This is a table.");

            var table = new Table(rows: 2, columns: 2) {Design = TableDesign.None};
            table.SetOutsideBorders(new Border(BorderStyle.Single, Uom.FromPoints(1))) //1pt
                 .SetInsideBorders(new Border(BorderStyle.Single, Uom.FromPoints(1.5))); // 1.5pt

            doc.Add(table);

            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);
            var rows = table.Rows.ToList();

            foreach (var row in rows)
            {
                for (int col = 0; col < row.ColumnCount; col++)
                {
                    var cell = row.Cells[col];
                    cell.Shading = new() {Fill = col % 2 == 0 ? Color.Pink : Color.LightBlue};

                    AddList(cell.Paragraphs.First(), 
                        nd, Enumerable.Range(1,5).Select(n => $"Item {n}"));    
                    
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
                paragraph.Text = text;
            }
        }
    }
}
