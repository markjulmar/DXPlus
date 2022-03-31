using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DXPlus;
using DXPlus.Charts;

namespace GenerateSampleDoc
{
    public static class Program
    {
        private static readonly List<Action<IDocument>> testers = new()
        {
            WriteTitle, 
            WriteFirstParagraph,
            AddCustomList,
            AddVideoToDoc,
            AddImageToDoc,
            AddHeaderAndFooter,
            CreateTableWithList,
            CreateNumberedList,
            CreateBulletedList,
            CreateBasicTable,
            AddTableToDocument,
            AddBookmarkToDocument,
            AddBarChartToDocument,
            AddLineChartToDocument,
            AddPieChartToDocument
        };

        public static void Main()
        {
            var fn = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.docx");
            using var doc = Document.Create(fn);

            testers.ForEach(f => f(doc));

            doc.Save();
            Console.WriteLine("Wrote document");
        }

        private static void WriteTitle(IDocument doc)
        {
            doc.AddRange(new[] { "Introduction", "This is some text" });
            doc.Paragraphs.First()
                .Style(HeadingType.Heading1)
                .InsertBefore(new Paragraph("This is a title").Style(HeadingType.Title))
                .InsertAfter(new Paragraph($"Last edited at {DateTime.Now.ToShortDateString()} by M. Smith").Style(HeadingType.Subtitle));
        }

        private static void WriteFirstParagraph(IDocument doc)
        {
            doc.Add("Highlighted text").Style(HeadingType.Heading2);
            doc.Add("First line. ")
                .AddText("This sentence is highlighted", new() { Highlight = Highlight.Yellow })
                .AddText(", but this is ")
                .AddText("not", new() { Italic = true })
                .AddText(".");

            doc.Add("This is another paragraph.")
                .Newline()
                .AddText("This is a second line. ")
                .AddText("It includes some ")
                .AddText("large", new Formatting { Font = new FontFamily("Times New Roman"), FontSize = 32 })
                .AddText(", blue", new Formatting { Color = Color.Blue })
                .AddText(", bold text.", new Formatting { Bold = true })
                .Newline()
                .AddText("And finally some normal text.");

            var paragraph = new Paragraph()
                .AddText("This line contains a ")
                .Add(new Hyperlink("link", new Uri("http://www.microsoft.com")))
                .AddText(". Here's a .");
            paragraph.Insert(paragraph.Text.Length-1, new Hyperlink("second link", new Uri("http://docs.microsoft.com/")));
            doc.Add(paragraph);

            // Simple equation
            doc.AddEquation("x = y+z");

            // Blue, large equation
            paragraph = new Paragraph();
            paragraph.AddEquation("x = (y+z)/t", new Formatting { FontSize = 18, Color = Color.Blue });
            doc.Add(paragraph);

            // Centered paragraph
            doc.Add(new Paragraph("I am centered 20pt Comic Sans.",
                    new() { Font = new FontFamily("Comic Sans MS"), FontSize = 20 })
                { Properties = new() { Alignment = Alignment.Center } });

            doc.Add(new Paragraph("This paragraph has the first sentence indented. "
                          + "It shows how you can use the Indent property to control how paragraphs are lined up.")
                    { Properties = new() { FirstLineIndent = Uom.FromInches(0.5) } })
                .Newline()
                .AddText("This line shouldn't be indented - instead, it should start over on the left side.");

            doc.AddPageBreak();
        }

        private static void AddVideoToDoc(IDocument doc)
        {
            doc.AddRange(new[] {
                "This is a video.",
                new Paragraph { Properties = new() {Alignment = Alignment.Center} }
                    .Add(doc.CreateVideo(
                        Path.Combine("images", "video-placeholder.png"),
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

        private static void AddImageToDoc(IDocument doc)
        {
            doc.AddParagraph();

            var svgImage = doc.CreateImage(Path.Combine("images", "test.svg"));
            var p = doc.Add("This is a picture:");
            p.Add(svgImage.CreatePicture(string.Empty, string.Empty));

            var comicImage = doc.CreateImage(Path.Combine("images", "1022.jpg"));
            var picture = comicImage.CreatePicture(string.Empty, string.Empty)
                .SetPictureShape(BasicShapes.Ellipse)
                .SetRotation(20)
                .SetDecorative(true)
                .SetName("Bat-Man!");

            doc.AddParagraph().Add(picture);
            picture.Drawing.AddCaption("The batman!");

            // Add with different size.
            p = doc.Add("And a final pic (dup of svg) sized to 200x200:");
            p.Add(svgImage.CreatePicture(200, 200));
        }

        private static void AddHeaderAndFooter(IDocument doc)
        {
            var mainSection = doc.Sections.First();

            var footer = mainSection.Footers.Default;
            footer.MainParagraph.Properties = new() {Alignment = Alignment.Right};
            footer.MainParagraph.Text = "Page ";
            footer.MainParagraph.AddPageNumber(PageNumberFormat.Normal);

            var image = doc.CreateImage(Path.Combine("images", "clock.png"));
            var picture = image.CreatePicture(48, 48);
            picture.IsDecorative = true;

            var header = mainSection.Headers.Default;
            header.MainParagraph.Text = "Welcome to the ";
            header.MainParagraph.Add(picture);
            header.MainParagraph.AddText(" tower!");
        }

        static void CreateNumberedList(IDocument doc)
        {
            doc.AddPageBreak();

            doc.Add("Numbered List").Style(HeadingType.Heading2);

            var numberStyle = doc.NumberingStyles.NumberStyle();
            doc.Add("First Item").ListStyle(numberStyle)
                .AddParagraph("First sub list item").ListStyle(numberStyle, level: 1)
                .AddParagraph("Second item.").ListStyle(numberStyle)
                .AddParagraph("Third item.").ListStyle(numberStyle)
                .AddParagraph("Nested item.").ListStyle(numberStyle, level: 1)
                .AddParagraph("Second nested item.").ListStyle(numberStyle, level: 1);
        }

        static void CreateBulletedList(IDocument doc)
        {
            const double fontSize = 15;
            doc.AddPageBreak();

            doc.Add("Lists with fonts").Style(HeadingType.Heading2);
            var style = doc.NumberingStyles.BulletStyle();

            foreach (var fontFamily in FontFamily.Families.Take(20))
            {
                doc.Add(new Paragraph(fontFamily.Name, 
                        new() {Font = fontFamily, FontSize = fontSize})
                    .ListStyle(style));
            }
        }

        static void AddCustomList(IDocument doc)
        {
            doc.AddPageBreak();

            var nd = doc.NumberingStyles.CustomBulletStyle("", new FontFamily("Wingdings"));
            nd.Style.Levels.First().Formatting.Color = Color.Green;

            doc.Add("Item #1").ListStyle(nd);
            doc.Add("Item #2").ListStyle(nd);
            doc.Add("Sub-item #1").ListStyle(nd, 1);
            doc.Add("Sub-item #2").ListStyle(nd, 1);
            doc.Add("Sub-item #1").ListStyle(nd, 2);
            doc.Add("Sub-item #2").ListStyle(nd, 2);
            doc.Add("Item #3").ListStyle(nd);
            doc.Add("Item #4").ListStyle(nd);
        }

        static void AddTableToDocument(IDocument doc)
        {
            doc.AddPageBreak();

            doc.Add("Large 10x10 table across whole page width")
                .Style(HeadingType.Heading2);

            var section = doc.Sections.First();
            Table table = new Table(10, 10);

            // Determine the width of the page
            double pageWidth = section.Properties.PageWidth - section.Properties.LeftMargin - section.Properties.RightMargin;
            double colWidth = pageWidth / table.ColumnCount;

            // Add some random data into the table
            foreach (var cell in table.Rows.SelectMany(row => row.Cells))
            {
                cell.Paragraphs.First().Text = new Random().Next().ToString();
                cell.CellWidth = colWidth;
                cell.SetMargins(0);
            }

            // Auto fit the table and set a border
            table.AutoFit();

            table.SetOutsideBorders(
                new Border(BorderStyle.DoubleWave, Uom.FromPoints(2)) { Color = Color.CornflowerBlue });

            // Insert the table into the document
            doc.Add("This line should be above the table.");
            var paragraph = doc.Add("... and this line below the table.");
            paragraph.InsertBefore(table);
        }

        static void CreateBasicTable(IDocument doc)
        {
            doc.AddPageBreak();
            doc.Add("Basic Table").Style(HeadingType.Heading2);

            var table = new Table(new[,]
            {
                { "Title", "# visitors" }, // header
                { "The wonderful world of Disney", "200,000,000 per year." },
                { "Star Wars experience", "1,000,000 per year." },
                { "Hogwarts", "10,000 per year." },
                { "Marvel town", "230,000 per year." }
            })
            {
                Design = TableDesign.ColorfulGridAccent6,
                ConditionalFormatting = TableConditionalFormatting.FirstRow,
                Alignment = Alignment.Center
            };

            doc.AddParagraph().InsertAfter(table);
        }

        static void CreateTableWithList(IDocument doc)
        {
            doc.AddPageBreak();
            doc.Add("This is a table.");

            var table = new Table(rows: 2, columns: 2) {Design = TableDesign.None};
            table.SetOutsideBorders(new Border(BorderStyle.Single, Uom.FromPoints(1))) //1pt
                 .SetInsideBorders(new Border(BorderStyle.Single, Uom.FromPoints(1.5))); // 1.5pt

            doc.Add(table);

            var nd = doc.NumberingStyles.NumberStyle();

            foreach (var row in table.Rows)
            {
                for (int col = 0; col < row.ColumnCount; col++)
                {
                    var cell = row.Cells[col];
                    cell.Shading = new() {Fill = col % 2 == 0 ? Color.Pink : Color.LightBlue};

                    int index = 0;
                    var paragraph = cell.Paragraphs.First();
                    foreach (var text in Enumerable.Range(1, 5).Select(n => $"Item {n}"))
                    {
                        if (index++ == 0)
                            paragraph.ListStyle(nd, 0);
                        else
                            paragraph = paragraph.AddParagraph().ListStyle(nd, 0);
                        paragraph.Text = text;
                    }
                }
            }

            table.AutoFit();
        }

        class CompanySales
        {
            public string Year { get; set; }
            // Sales in millions of units
            public double TotalSales { get; set; }

            public static CompanySales[] Acme
            {
                get
                {
                    return new[]
                    {
                        new CompanySales {Year = "2016", TotalSales = 1.2},
                        new CompanySales {Year = "2017", TotalSales = 2.4},
                        new CompanySales {Year = "2018", TotalSales = 3.6},
                        new CompanySales {Year = "2019", TotalSales = 5.8}
                    };
                }
            }

            public static CompanySales[] Cyberdyne
            {
                get
                {
                    return new[]
                    {
                        new CompanySales {Year = "2016", TotalSales = 10.5},
                        new CompanySales {Year = "2017", TotalSales = 11.9},
                        new CompanySales {Year = "2018", TotalSales = 16.6},
                        new CompanySales {Year = "2019", TotalSales = 25.3}
                    };
                }
            }

        }


        static void AddBarChartToDocument(IDocument doc)
        {
            doc.AddPageBreak();

            var chart = new BarChart();
            chart.AddLegend(ChartLegendPosition.Bottom, false);

            // Add series
            var acmeSeries = new Series("ACME") { Color = Color.DarkBlue };
            acmeSeries.Bind(CompanySales.Acme, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
            chart.AddSeries(acmeSeries);

            var sprocketsSeries = new Series("Cyberdyne") { Color = Color.FromArgb(1, 0xff, 0, 0xff) };
            sprocketsSeries.Bind(CompanySales.Cyberdyne, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
            chart.AddSeries(sprocketsSeries);

            // Insert chart into document
            doc.AddParagraph().Add("sales", chart);
        }

        static void AddLineChartToDocument(IDocument doc)
        {
            doc.AddPageBreak();

            var chart = new LineChart();
            chart.AddLegend(ChartLegendPosition.Bottom, false);

            // Add series
            var acmeSeries = new Series("ACME") { Color = Color.DarkBlue };
            acmeSeries.Bind(CompanySales.Acme, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
            chart.AddSeries(acmeSeries);

            var sprocketsSeries = new Series("Cyberdyne") { Color = Color.FromArgb(1, 0xff, 0, 0xff) };
            sprocketsSeries.Bind(CompanySales.Cyberdyne, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
            chart.AddSeries(sprocketsSeries);

            // Insert chart into document
            doc.AddParagraph().Add("sales2", chart);
        }

        static void AddPieChartToDocument(IDocument doc)
        {
            doc.AddPageBreak();

            var chart = new PieChart();
            chart.AddLegend(ChartLegendPosition.Bottom, false);

            // Add series
            var acmeSeries = new Series("ACME");
            acmeSeries.Bind(CompanySales.Acme, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
            chart.AddSeries(acmeSeries);

            // Insert chart into document
            doc.AddParagraph().Add("sales3", chart);
        }

        static void AddBookmarkToDocument(IDocument doc)
        {
            doc.AddPageBreak();
            doc.Add("Bookmark test").Style(HeadingType.Heading2);

            // Add a bookmark
            var paragraph = doc
                .Add("This is a paragraph which contains a ")
                .AddBookmark("namedBookmark")
                .AddText(" bookmark");

            // Set text at the bookmark
            paragraph.InsertTextAtBookmark("namedBookmark", "handy ");
        }
    }
}
