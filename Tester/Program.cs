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

            //AddImageToDoc(doc);
            //CreateDocWithHeaderAndFooter(doc);
            //WriteTitle(doc);
            AddVideoToDoc(doc);
            
            doc.Save();
            Console.WriteLine("Wrote document");

            CheckDocument(fn);
        }

        private static void CheckDocument(string fn)
        {
            using var doc = Document.Load(fn);

            var section = doc.Sections.First();
            Console.WriteLine(section.Headers.Default.MainParagraph);
        }

        private static void AddVideoToDoc(IDocument doc)
        {
            doc.Add("This is an introduction.");
            
            var p = doc.AddParagraph();
            p.Properties.Alignment = Alignment.Center;

            var image = doc.CreateVideo(
                "video-placeholder.png",
                new Uri("https://www.microsoft.com/en-us/videoplayer/embed/RWwMdr", UriKind.Absolute),
                400, 225);

            if (image != null)
                p.Append(image);
            
            doc.Add("And a closing paragraph.");
        }

        private static void WriteTitle(IDocument doc)
        {
            doc.Add("Introduction").Style(HeadingType.Heading1);
            doc.Add("This is some text");

            var p = doc.Paragraphs.First();
            p.InsertBefore(new Paragraph("This is a title").Style(HeadingType.Title))
                .Append(new Paragraph($"Last edited at {DateTime.Now.ToShortDateString()} by M. Smith").Style(
                    HeadingType.Subtitle));
            
            doc.Save();
        }

        private static void AddImageToDoc(IDocument doc)
        {
            var img = doc.CreateImage(@"test.svg");
            var p = doc.Add("This is a picture:");
            p.Append(img.CreatePicture(string.Empty, string.Empty));

            var im2 = doc.CreateImage(@"test2.png");
            p.Append(im2.CreatePicture(string.Empty, string.Empty));

            // Add with different size.
            p = doc.Add("And a final pic (dup of svg!):");
            p.Append(img.CreatePicture(50, 50));
        }

        private static void CreateDocWithHeaderAndFooter(IDocument doc)
        {
            var mainSection = doc.Sections.First();
            var paragraphs = mainSection.Paragraphs.ToList();
            var header = mainSection.Headers.Default;

            var p1 = header.MainParagraph;
            p1.SetText("This is some text - ");
            p1.AddPageNumber(PageNumberFormat.Normal);

            doc.Add("This is the first paragraph");
            doc.AddPageBreak();
            doc.Add("This is page 2");
        }

        static void CreateTableWithList(IDocument doc)
        {
            int rows = 2;
            int columns = 2;

            var documentTable = doc.AddTable(rows: rows, columns: columns);
            var nd = doc.NumberingStyles.Create(NumberingFormat.Bullet);

            documentTable.Design = TableDesign.Normal;

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
