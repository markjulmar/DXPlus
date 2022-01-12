using System;
using System.IO;
using System.Linq;
using DXPlus;

namespace Tester
{
    class Program
    {
        public static void Main(string[] args)
        {
            string fn = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.docx");
            using var doc = Document.Create(fn);

            AddImageToDoc(doc);
            //CreateDocWithHeaderAndFooter(doc);
            //WriteTitle(doc);
            //AddVideoToDoc(doc);
            
            doc.Save();
            Console.WriteLine("Wrote document");
        }

        private static void AddVideoToDoc(IDocument doc)
        {
            doc.AddParagraph("This is an introduction.");
            
            var p = doc.AddParagraph();
            p.Properties.Alignment = Alignment.Center;

            p.Append(doc.CreateVideo(
                "video-placeholder.png",
                new Uri("https://www.microsoft.com/en-us/videoplayer/embed/RWwMdr", UriKind.Absolute),
                400, 225));
            
            doc.AddParagraph("And a closing paragraph.");
        }

        private static void WriteTitle(IDocument doc)
        {
            doc.AddParagraph("Introduction").Style(HeadingType.Heading1);
            doc.AddParagraph("This is some text");

            var p = doc.Paragraphs.First();
            p.InsertBefore(new Paragraph("This is a title").Style(HeadingType.Title))
                .Append(new Paragraph($"Last edited at {DateTime.Now.ToShortDateString()} by M. Smith").Style(
                    HeadingType.Subtitle));
            
            doc.Save();
        }

        private static void AddImageToDoc(IDocument doc)
        {
            var img = doc.AddImage(@"test.svg");
            var p = doc.AddParagraph("This is a picture:");
            p.Append(img.CreatePicture());

            var im2 = doc.AddImage(@"test2.png");
            p.Append(im2.CreatePicture());

            // Add with different size.
            p = doc.AddParagraph("And a final pic (dup of svg!):");
            p.Append(img.CreatePicture(50, 50));
        }

        private static void CreateDocWithHeaderAndFooter(IDocument doc)
        {
            var mainSection = doc.Sections.First();
            var header = mainSection.Headers.Default;

            var p1 = header.MainParagraph;
            p1.SetText("This is some text - ");
            p1.AddPageNumber(PageNumberFormat.Normal);

            doc.AddParagraph("This is the firt paragraph");
            doc.AddPageBreak();
            doc.AddParagraph("This is page 2");
        }

        static void CreateTableWithList(IDocument doc)
        {
            int rows = 2;
            int columns = 2;

            var documentTable = doc.AddTable(rows: rows, columns: columns);
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            documentTable.Design = TableDesign.TableNormal;

            AddListToCell(
                documentTable.Rows.ElementAt(1).Cells[0], nd,
                new[] { "Item 1", "Item 2", "Item 3", "Item 4" });

            AddListToCell(
                documentTable.Rows.ElementAt(1).Cells[1], nd,
                new[] { "Item 5", "Item 6", "Item 7", "Item 8" });

            documentTable.AutoFit = true;

            var t = doc.AddTable(3, 3);
        }

        private static void AddListToCell(TableCell cell, NumberingDefinition nd, string[] terms)
        {
            var cellParagraph = cell.Paragraphs.First();
            for (int i = 0; i < terms.Length; i++)
            {
                string text = terms[i];
                if (i == 0)
                    cellParagraph.ListStyle(nd, 0);
                else
                    cellParagraph = cellParagraph.AddParagraph().ListStyle(nd,0);
                cellParagraph.SetText(text);
            }
        }
    }
}
