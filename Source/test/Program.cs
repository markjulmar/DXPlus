using DXPlus;
using DXPlus.Charts;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace TestDXPlus
{
    public static class Program
    {
        private static void Main()
        {
            const string testFolder = "docs";
            if (Directory.Exists(testFolder))
                Directory.Delete(testFolder, true);
            Directory.CreateDirectory(testFolder);
            Directory.SetCurrentDirectory(testFolder);

            var document = Document.Create("testDocument.docx");

            AddHeader(document);

            // Add a title.
            document.AddParagraph("Welcome")
                .Heading(HeadingType.Title)
                .AppendLine();

            AddBasicText(document);
            AddPicture(document);
            AddRTLText(document);
            AddIndentedParagraph(document);
            AddHyperlinks(document);
            AddLists(document);
            AddEquations(document);
            AddBookmarks(document);
            AddCharts(document);
            AddTables(document);
            ShowAllHeaderStyles(document);

            //// Add ToC
            document.InsertTableOfContents("Table of Contents",
                TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading1");

            document.Save();
        }

        private static void AddHeader(IDocument document)
        {
            // Add a first page header
            var section = document.Sections.Single();
            section.Headers.First
                .Add().Append("First page header").Bold();

            // Add an image into the document.    
            var image = document.AddImage(Path.Combine("..", "images", "logo_the_happy_builder.png"));

            // Create a picture and add it to the document.
            Picture picture = image.CreatePicture();

            section.Headers.Default.Add()
                .Append(picture);
        }

        private static void ShowAllHeaderStyles(IDocument document)
        {
            document.AddPageBreak();
            document.AddParagraph("All the header styles")
                .Heading(HeadingType.Heading1)
                .AppendLine();

            foreach (var heading in (HeadingType[]) Enum.GetValues(typeof(HeadingType)))
            {
                document.AddParagraph($"{heading} - The quick brown fox jumps over the lazy dog")
                    .Heading(heading);
            }
        }

        private static void AddTables(IDocument document)
        {
            document.AddPageBreak();
            document.AddParagraph("Tables").Heading(HeadingType.Heading1).AppendLine();

            document.AddParagraph("Basic Table").Heading(HeadingType.Heading2).AppendLine();

            var table = new Table(new[,] {{"Title", "The wonderful world of Disney"}, {"# visitors", "200,000,000 per year."}})
            {
                Design = TableDesign.ColorfulList,
                Alignment = Alignment.Center
            };

            document.AddParagraph()
                    .AddTable(table);

            document.AddParagraph("2x2 table inserted into middle of paragraph").Heading(HeadingType.Heading2).AppendLine(); ;
            table = new Table(2, 2) {Design = TableDesign.MediumGrid1Accent1, Alignment = Alignment.Center};

            table.Rows[0].Cells[0].Paragraphs[0].Append("One");
            table.Rows[0].Cells[1].Paragraphs[0].Append("Two");
            table.Rows[1].Cells[0].Paragraphs[0].Append("Three");
            table.Rows[1].Cells[1].Paragraphs[0].Append("Four");

            Paragraph p;
            document.AddParagraph("This line should be above the table.");
            p = document.AddParagraph().AppendLine("... and this line below the table.");
            p.InsertTableBefore(table);
            p.AppendLine();

            AddLargeTable(document);

            document.AddParagraph("Table with merged/centered cells").Heading(HeadingType.Heading2).AppendLine(); ;

            table = new Table(new[,] {{"1", "2", "3"}, {"4", "5", "6"}, {"7", "8", "9"}, {"10", "11", "12"},});

            table.Rows[1].MergeCells(0, table.ColumnCount);
            table.Rows[1].Cells.SelectMany(c => c.Paragraphs).ToList().ForEach(p => p.Alignment(Alignment.Center));
            document.AddTable(table);

            document.AddParagraph("Table with merged/centered rows").Heading(HeadingType.Heading2).AppendLine();

            table = new Table(new[,]
            {
                {"1", "2", "3"},
                {"4", "5", "6"},
                {"7", "8", "9"},
                {"10", "11", "12"},
            });

            document.AddTable(table);
            table.MergeCellsInColumn(1, 0, table.Rows.Count);

            document.AddParagraph("Empty 2x1 table").Heading(HeadingType.Heading2).AppendLine();
            document.AddTable(new Table(2, 1));
        }

        private static void AddCharts(IDocument document)
        {
            document.AddPageBreak();
            document.AddParagraph("Charts").Heading(HeadingType.Heading1).AppendLine();
            BarChart(document);
            PieChart(document);
            LineChart(document);
            Chart3D(document);
        }

        private static void AddBookmarks(IDocument document)
        {
            document.AddParagraph("Bookmarks").Heading(HeadingType.Heading1).AppendLine();
            Paragraph p;
            p = document.AddParagraph("This is a paragraph which contains a ")
                .AppendBookmark("secondBookmark").Append("bookmark");
            p.InsertAtBookmark("secondBookmark", "handy ");
        }

        private static void AddEquations(IDocument document)
        {
            document.AddParagraph("Equations").Heading(HeadingType.Heading1).AppendLine();
            document.AddEquation("x = y+z");
            document.AddEquation("x = (y+z)/t").FontSize(18).Color(Color.Blue);
        }

        private static void AddLists(IDocument document)
        {
            // Add two lists.
            document.AddParagraph("Lists").Heading(HeadingType.Heading1).AppendLine();
            document.AddParagraph("Numbered List").Heading(HeadingType.Heading2).AppendLine();

            List numberedList = new List(NumberingFormat.Numbered)
                .AddItem("First item.")
                .AddItem("First sub list item", level: 1)
                .AddItem("Second item.")
                .AddItem("Third item.")
                .AddItem("Nested item.", level: 1)
                .AddItem("Second nested item.", level: 1);
            document.AddList(numberedList);

            document.AddParagraph("Bullet List").Heading(HeadingType.Heading2).AppendLine();
            List bulletedList = new List(NumberingFormat.Bulleted)
                .AddItem("First item.")
                .AddItem("Second item")
                .AddItem("Sub bullet item", level: 1)
                .AddItem("Second sub bullet item", level: 1)
                .AddItem("Third item");
            document.AddList(bulletedList);

            document.AddParagraph("Lists with fonts").Heading(HeadingType.Heading2).AppendLine();
            foreach (var fontFamily in FontFamily.Families.Take(5))
            {
                const double fontSize = 15;
                bulletedList = new List(NumberingFormat.Bulleted) { Font = fontFamily, FontSize = fontSize }
                    .AddItem("One")
                    .AddItem("Two (L1)", level: 1)
                    .AddItem("Three (L2)", level: 2)
                    .AddItem("Four");
                document.AddList(bulletedList);
            }
        }

        private static void AddHyperlinks(IDocument document)
        {
            // Add two hyperlinks to the document.
            Paragraph p;
            document.AddParagraph("Hyperlinks").Heading(HeadingType.Heading2).AppendLine();
            p = document.AddParagraph()
                .AppendLine("This line contains a ")
                .Append(new Hyperlink("link", new Uri("http://www.microsoft.com")))
                .Append(". With a few lines of text to read.")
                .AppendLine("And a final line with a .");
            p.InsertHyperlink(new Hyperlink("second link", new Uri("http://docs.microsoft.com/")), p.Text.Length - 1);
        }

        private static void AddIndentedParagraph(IDocument document)
        {
            // Indentation.
            document.AddParagraph("Indented text").Heading(HeadingType.Heading2).AppendLine();
            document.AddParagraph("This paragraph has the first sentence indented. "
                                  + "It shows how you can use the Intent property to control how paragraphs are lined up.")
                .FirstLineIndent(20)
                .AppendLine("This line shouldn't be indented - instead, it should start over on the left side.");
        }

        private static void AddRTLText(IDocument document)
        {
            // Try RTL text
            Paragraph p;
            document.AddParagraph("RTL text").Heading(HeadingType.Heading2).AppendLine();
            p = document.AddParagraph("Hello World - RightToLeft (should be RTL).");
            p.Direction = Direction.RightToLeft;
        }

        private static void AddPicture(IDocument document)
        {
            // Add an image into the document.    
            var image = document.AddImage(Path.Combine("..", "images", "logo_template.png"));

            // Create a picture and add it to the document.
            Picture picture = image.CreatePicture()
                .SetPictureShape(BasicShapes.Ellipse)
                .SetRotation(45);

            // Insert a new Paragraph into the document.
            document.AddParagraph("Pictures").Heading(HeadingType.Heading2).AppendLine();
            document.AddParagraph()
                .AppendLine("Just below there should be a picture rotated 10 degrees.")
                .Append(picture)
                .AppendLine();

            // Add a second copy of the same image
            document.AddParagraph()
                .AppendLine("Lets add another picture (without the fancy  rotation stuff)")
                .Append(image.CreatePicture());
        }

        private static void AddBasicText(IDocument document)
        {
            // Start with some hello world text.
            document.AddParagraph("Hello, World! This is the first paragraph.")
                .AppendLine()
                .Append("It includes some ")
                .Append("large")
                .Font(new FontFamily("Times New Roman"))
                .FontSize(32)
                .Append(", blue")
                .Color(Color.Blue)
                .Append(", bold text.")
                .Bold()
                .AppendLine()
                .AppendLine("And finally some normal text.");

            // Use a paragraph object directly.
            var p = document.AddParagraph().Append("Styled Text").Heading(HeadingType.Heading2).AppendLine();

            p.Append("I am ").Append("bold").Bold()
                .Append(" and I am ")
                .Append("italic").Italic().Append(".")
                .AppendLine("I am ")
                .Append("Arial Black")
                .Font(new FontFamily("Arial Black"))
                .Append(" and I am not.")
                .AppendLine("I am centered 20pt Comic Sans")
                .FontSize(20)
                .Font(new FontFamily("Comic Sans MS"))
                .Alignment(Alignment.Center).Append(" and I am not.")
                .AppendLine("I am ")
                .Append("Blue").Color(Color.Blue)
                .Append(" and I am ")
                .Append("Red").Color(Color.Red).Append(".");

            // Try some highlighted words
            document.AddParagraph("Highlighted text").Heading(HeadingType.Heading2).AppendLine();
            document.AddParagraph("First line. ")
                .Append("This sentence is highlighted")
                .Highlight(Highlight.Yellow)
                .Append(", but this is ")
                .Append("not").Italic()
                .Append(".");
        }

        private static void BarChart(IDocument document)
        {
            document.AddParagraph("Bar Chart").Heading(HeadingType.Heading2).AppendLine();

            // Create chart.
            var chart = new BarChart
            {
                BarDirection = BarDirection.Column,
                BarGrouping = BarGrouping.Standard,
                GapWidth = 400
            };

            chart.AddLegend(ChartLegendPosition.Bottom, false);

            // Create data.
            var company1 = ChartData.CreateCompanyList1();
            var company2 = ChartData.CreateCompanyList2();

            // Create and add series
            var series1 = new Series("Microsoft") { Color = Color.DarkBlue };
            series1.Bind(company1, nameof(ChartData.Month), nameof(ChartData.Money));
            chart.AddSeries(series1);

            var series2 = new Series("Apple") { Color = Color.FromArgb(1, 0xff, 0, 0xff)};
            series2.Bind(company2, nameof(ChartData.Month), nameof(ChartData.Money));
            chart.AddSeries(series2);

            // Insert chart into document
            document.InsertChart(chart);
        }

        private static void Chart3D(IDocument document)
        {
            document.AddParagraph("3D Chart").Heading(HeadingType.Heading2).AppendLine();

            var company1 = ChartData.CreateCompanyList1();
            var series = new Series("Microsoft") {Color = Color.GreenYellow};
            series.Bind(company1, nameof(ChartData.Month), nameof(ChartData.Money));

            var barChart = new BarChart { View3D = true };
            barChart.AddSeries(series);

            // Insert chart into document
            document.InsertChart(barChart);
        }

        private static void AddLargeTable(IDocument document)
        {
            document.AddParagraph("Large 10x10 Table across whole page width").Heading(HeadingType.Heading2).AppendLine();

            var section = document.Sections.Last();
            Table table = document.AddTable(10,10);

            double pageWidth = section.Properties.PageWidth - section.Properties.LeftMargin - section.Properties.RightMargin;
            double colWidth = pageWidth / table.ColumnCount;

            foreach (var cell in table.Rows.SelectMany(row => row.Cells))
            {
                cell.Paragraphs[0].SetText(new Random().Next().ToString());
                cell.Width = colWidth;
                cell.SetMargins(0);
            }

            table.AutoFit = AutoFit.Contents;
            TableBorder border = new TableBorder(TableBorderStyle.DoubleWave, 0, 0, Color.CornflowerBlue);
            table.SetBorders(border);
        }

        private static void LineChart(IDocument document)
        {
            document.AddParagraph("Line Chart").Heading(HeadingType.Heading2).AppendLine();

            // Create chart.
            LineChart c = new LineChart();
            c.AddLegend(ChartLegendPosition.Bottom, false);

            // Create data.
            List<ChartData> company1 = ChartData.CreateCompanyList1();
            List<ChartData> company2 = ChartData.CreateCompanyList2();

            // Create and add series
            Series s1 = new Series("Microsoft")
            {
                Color = Color.GreenYellow
            };
            s1.Bind(company1, nameof(ChartData.Month), nameof(ChartData.Money));
            c.AddSeries(s1);

            Series s2 = new Series("Apple");
            s2.Bind(company2, nameof(ChartData.Month), nameof(ChartData.Money));
            c.AddSeries(s2);

            // Insert chart into document
            document.InsertChart(c);
        }

        private static void PieChart(IDocument document)
        {
            document.AddParagraph("Pie Chart").Heading(HeadingType.Heading2).AppendLine();

            // Create chart.
            PieChart c = new PieChart();
            c.AddLegend(ChartLegendPosition.Bottom, false);

            // Create data.
            List<ChartData> company2 = ChartData.CreateCompanyList2();

            // Create and add series
            Series s = new Series("Apple");
            s.Bind(company2, nameof(ChartData.Month), nameof(ChartData.Money));
            c.AddSeries(s);

            // Insert chart into document
            document.InsertChart(c);
        }

        private class ChartData
        {
            public double Money { get; set; }
            public string Month { get; set; }
            public static List<ChartData> CreateCompanyList1()
            {
                return new List<ChartData>
                {
                    new ChartData() { Month = "January", Money = 100 },
                    new ChartData() { Month = "February", Money = 120 },
                    new ChartData() { Month = "March", Money = 140 }
                };
            }

            public static List<ChartData> CreateCompanyList2()
            {
                return new List<ChartData>
                {
                    new ChartData() { Month = "January", Money = 80 },
                    new ChartData() { Month = "February", Money = 160 },
                    new ChartData() { Month = "March", Money = 130 }
                };
            }
        }
    }
}
