using DXPlus;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using DXPlus.Charts;

namespace TestDXPlus
{
    public static class Program
    {
        private static void Main()
        {
            Setup("docs");

            HelloWorld();
            HighlightWords();
            HelloWorldAdvancedFormatting();
            HelloWorldProtectedDocument();
            HelloWorldAddPictureToWord();
            RightToLeft();
            Indentation();
            FirstPageHeader();
            HyperlinksInDocument();
            AddList();
            Equations();
            Bookmarks();
            BookmarksReplaceTextOfBookmarkKeepingFormat();
            BarChart();
            PieChart();
            LineChart();
            Chart3D();
            DocumentMargins();
            CreateTableWithTextDirection();
            AddToc();
            AddTocByReference();
            TablesDocument();
            DocumentsWithListsFontChange();
            DocumentHeading();
            LargeTable();
            ProgrammaticallyManipulateImbeddedImage();
            CountNumberOfParagraphs();
            EmptyTable();
            MergeTableRows();
            MergeTableColumns();
        }

        private static void MergeTableColumns()
        {
            Enter();

            var doc = Document.Create("mergeTableColumns.docx");

            doc.AddParagraph("Check out the table below.")
               .Heading(HeadingType.Heading2)
               .AppendLine();

            var t = new Table(new[,] {{"1", "2", "3"}, {"4", "5", "6"}, {"7", "8", "9"}, {"10", "11", "12"},})
            {
                Alignment = Alignment.Center
            };

            t.Rows[1].MergeCells(0, t.ColumnCount);
            t.Rows[1].Cells[0].VerticalAlignment = VerticalAlignment.Center;

            doc.AddTable(t);
            doc.Save();
        }

        private static void MergeTableRows()
        {
            Enter();

            var doc = Document.Create("mergeTableRow.docx");

            doc.AddParagraph("Check out the table below.");

            var t = new Table(new[,]
            {
                { "1", "2", "3" },
                { "4", "5", "6" },
                { "7", "8", "9" },
                { "10", "11", "12" },
            });

            doc.AddTable(t);

            t.MergeCellsInColumn(1, 0, t.Rows.Count);
            t.Alignment = Alignment.Center;
            doc.Save();
        }

        private static void EmptyTable()
        {
            Enter();

            var doc = Document.Create("emptyTable.docx");
            doc.AddTable(new Table(2, 1));
            doc.Save();
        }

        private static void CountNumberOfParagraphs()
        {
            Enter();

            var doc = Document.Load(Path.Combine("..", "Input.docx"));

            foreach (var p in doc.Paragraphs.Where(p => !string.IsNullOrEmpty(p.Text)))
            {
                Console.WriteLine(p.Text);
            }
        }

        private static void Enter([CallerMemberName] string name = "")
        {
            Console.WriteLine($"=> {name}");
        }

        private static void DocumentHeading()
        {
            Enter();

            var document = Document.Create("documentHeading.docx");

            foreach (var heading in (HeadingType[])Enum.GetValues(typeof(HeadingType)))
            {
                document.AddParagraph($"{heading} - The quick brown fox jumps over the lazy dog")
                        .Heading(heading);
            }

            document.Save();
        }

        private static void AddList()
        {
            Enter();

            var document = Document.Create("Lists.docx");

            document.AddParagraph("Numbered List")
                .Heading(HeadingType.Heading1);

            List numberedList = new List(ListItemType.Numbered)
                                        .AddItem("First List Item.")
                                        .AddItem("First sub list item", level: 1)
                                        .AddItem("Second List Item.")
                                        .AddItem("Third list item.")
                                        .AddItem("Nested item.", level: 1)
                                        .AddItem("Second nested item.", level: 1);
            document.AddList(numberedList);

            document.AddParagraph("Bullet List")
                    .Heading(HeadingType.Heading1);

            List bulletedList = new List(ListItemType.Bulleted)
                                        .AddItem("First Bulleted Item.")
                                        .AddItem("Second bullet item")
                                        .AddItem("Sub bullet item", level: 1)
                                        .AddItem("Second sub bullet item", level: 1)
                                        .AddItem("Third bullet item");
            document.AddList(bulletedList);

            foreach (var list in document.Lists)
            {
                Console.WriteLine($"{list.ListType} List {list.NumId} starting at {list.StartNumber}");
                foreach (var item in list.Items)
                {
                    Console.WriteLine($"\t{item.IndentLevel}> {item.Paragraph.Text}");
                }
            }

            document.Save();
        }

        private static void AddToc()
        {
            Enter();

            var document = Document.Create("Toc.docx");

            document.AddParagraph("Welcome to the document").Heading(HeadingType.Title).Align(Alignment.Center);
            document.AddPageBreak();

            document.InsertTableOfContents("Table of Contents",
                TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2");

            document.AddPageBreak();
            document.AddParagraph("Page #1").Heading(HeadingType.Heading1);
            document.AddParagraph("Some very interesting content here");
            document.AddParagraph("Heading 2").Style("Heading2");

            document.AddPageBreak();
            document.AddParagraph("Page #2").Heading(HeadingType.Heading1);
            document.AddParagraph("Some very interesting content here as well");
            document.AddParagraph("Heading 3").Style("Heading3");
            document.AddParagraph("Not so very interesting....");

            document.Save();
        }

        private static void AddTocByReference()
        {
            Enter();

            var document = Document.Create("TocByReference.docx");

            document.AddParagraph("Heading 1").Style("Heading1");
            document.AddParagraph("Some very interesting content here");
            document.AddParagraph("Heading 2").Style("Heading1");
            document.AddPageBreak();

            document.AddParagraph("Some very interesting content here as well");
            Paragraph h2 = document.AddParagraph("Heading 2.1").Style("Heading2");
            document.AddParagraph("Not so very interesting....");

            document.InsertTableOfContents(h2, "Table of Contents Goes Here",
                TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2");

            document.Save();
        }

        private static void BarChart()
        {
            Enter();

            var document = Document.Create("BarChart.docx");

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
            document.AddParagraph("Diagram").FontSize(20);
            document.InsertChart(chart);

            document.Save();
        }

        private static void Bookmarks()
        {
            Enter();

            var document = Document.Create("Bookmarks.docx");

            document.AddBookmark("firstBookmark");

            var paragraph2 = document.AddParagraph("This is a paragraph which contains a ")
                                     .AppendBookmark("secondBookmark").Append("bookmark");

            paragraph2.InsertAtBookmark("secondBookmark", "handy ");

            document.Save();
        }

        /// <summary>
        /// Loads a document 'DocumentWithBookmarks.docx' and changes text inside bookmark keeping formatting the same.
        /// This code creates the file 'BookmarksReplaceTextOfBookmarkKeepingFormat.docx'.
        /// </summary>
        private static void BookmarksReplaceTextOfBookmarkKeepingFormat()
        {
            Enter();

            var docX = Document.Load(Path.Combine("..", "DocumentWithBookmarks.docx"));

            foreach (var bookmark in docX.Bookmarks)
            {
                Console.WriteLine("Found bookmark {0}", bookmark.Name);
            }

            // Replace bookmarks content
            docX.Bookmarks["bmkNoContent"].SetText("Here there was a bookmark");
            docX.Bookmarks["bmkContent"].SetText("Here there was a bookmark with a previous content");
            docX.Bookmarks["bmkFormattedContent"].SetText("Here there was a formatted bookmark");

            docX.SaveAs("BookmarksReplaceTextOfBookmarkKeepingFormat.docx");
        }

        private static void Chart3D()
        {
            Enter();

            var document = Document.Create("3DChart.docx");

            var company1 = ChartData.CreateCompanyList1();
            var series = new Series("Microsoft") {Color = Color.GreenYellow};
            series.Bind(company1, nameof(ChartData.Month), nameof(ChartData.Money));

            var barChart = new BarChart { View3D = true };
            barChart.AddSeries(series);

            // Insert chart into document
            document.AddParagraph("3D Diagram").FontSize(20);
            document.InsertChart(barChart);

            document.Save();
        }

        private static void CreateTableWithTextDirection()
        {
            Enter();

            var document = Document.Create("createTableWithTextDirection.docx");

            var t = new Table(new[,]
            {
                { "A", "B", "C"},
                { "D", "E", "F"}

            }) {Alignment = Alignment.Left, Design = TableDesign.MediumGrid1Accent2};

            foreach (var cell in t.Rows[0].Cells)
                cell.TextDirection = TextDirection.BottomToTopLeftToEnd;

            document.AddTable(t);
            document.Save();
        }

        private static void DocumentMargins()
        {
            Enter();

            var document = Document.Create("DocumentMargins.docx");

            // Create a float var that contains doc Margins properties.
            double leftMargin = document.MarginLeft;
            double rightMargin = document.MarginRight;
            double topMargin = document.MarginTop;
            double bottomMargin = document.MarginBottom;

            Console.WriteLine($"Current margins: L={leftMargin}, R={rightMargin}, T={topMargin}, B={bottomMargin}");

            leftMargin = 95;
            rightMargin = 45;
            topMargin = 50;
            bottomMargin = 180;

            // Or simply work the margins by setting the property directly. 
            document.MarginLeft = leftMargin;
            document.MarginRight = rightMargin;
            document.MarginTop = topMargin;
            document.MarginBottom = bottomMargin;

            var bulletedList = new List(ListItemType.Bulleted)
                .AddItem("First Bulleted Item.")
                .AddItem("Second bullet item")
                .AddItem("Sub bullet item", level: 1)
                .AddItem("Second sub bullet item", level: 1)
                .AddItem("Third bullet item");
            document.AddList(bulletedList);

            // Save this document.
            document.Save();
        }

        private static void DocumentsWithListsFontChange()
        {
            Enter();

            var document = Document.Create("DocumentsWithListsFontChange.docx");

            foreach (var oneFontFamily in FontFamily.Families)
            {
                FontFamily fontFamily = oneFontFamily;
                const double fontSize = 15;

                var numberedList = new List(ListItemType.Numbered)
                    .AddItem("First List Item.")
                    .AddItem("First sub list item", level: 1)
                    .AddItem("Second List Item.")
                    .AddItem("Third list item.")
                    .AddItem("Nested item.", level: 1)
                    .AddItem("Second nested item.", level: 1);
                document.AddList(numberedList, fontFamily, fontSize);

                var bulletedList = new List(ListItemType.Bulleted)
                    .AddItem("First Bulleted Item.")
                    .AddItem("Second bullet item")
                    .AddItem("Sub bullet item", level: 1)
                    .AddItem("Second sub bullet item", level: 1)
                    .AddItem("Third bullet item");
                document.AddList(bulletedList);
            }
            
            document.Save();
        }

        /// <summary>
        /// Create a document with two equations.
        /// </summary>
        private static void Equations()
        {
            Enter();

            var document = Document.Create("Equations.docx");

            document.AddEquation("x = y+z");
            document.AddEquation("x = (y+z)/t").FontSize(18).Color(Color.Blue);

            document.Save();
        }

        private static void FirstPageHeader()
        {
            Enter();

            var document = Document.Create("firstPageHeader.docx");

            document.Headers.First
                    .Add().Append("First page header").Bold();

            document.AddParagraph("This is page #1");

            // Save all changes to this document.
            document.Save();
        }

        /// <summary>
        /// Creates a simple document with the text Hello World.
        /// </summary>
        private static void HelloWorld()
        {
            Enter();

            var document = Document.Create("HelloWorld.docx");
            Paragraph p = document.AddParagraph("Hello, World!").AppendLine();

            // Append some text and add formatting.
            p.Append("Some ")
                .Append("big")
                .Font(new FontFamily("Times New Roman"))
                .FontSize(32)
                .Append(", blue")
                .Color(Color.Blue)
                .Append(", bold text.")
                .Bold();

            p.AppendLine().Bold();
            p.AppendLine("This is some normal text.");

            document.Save();
        }

        /// <summary>
        /// Create a document with two pictures. One picture is inserted normal way, the other one with rotation
        /// </summary>
        private static void HelloWorldAddPictureToWord()
        {
            Enter();

            var document = Document.Create("HelloWorldAddPictureToWord.docx");

            // Add an image into the document.    
            DXPlus.Image image = document.AddImage(Path.Combine("..", "images", "logo_template.png"));

            // Create a picture (A custom view of an Image).
            Picture picture = image.CreatePicture()
                                   .SetPictureShape(BasicShapes.Ellipse);
            picture.Rotation = 10;

            // Insert a new Paragraph into the document.
            Paragraph title = document.AddParagraph()
                .Append("This is a test for a picture")
                .FontSize(20)
                .Font(new FontFamily("Comic Sans MS"));
            title.Alignment = Alignment.Center;

            // Insert a new Paragraph into the document.
            Paragraph p1 = document.AddParagraph();

            // Append content to the Paragraph
            p1.AppendLine("Just below there should be a picture ")
              .Append("picture").Bold()
              .Append(" inserted in a non-conventional way.")
              .AppendLine()
              .AppendLine("Check out this picture ")
              .AppendPicture(picture)
              .Append(" its funky don't you think?")
              .AppendLine();

            // Insert a new Paragraph into the document.
            document.AddParagraph()
                    .AppendLine("Is it correct?")
                    .AppendLine();

            // Lets add another copy of the image
            Picture pictureNormal = image.CreatePicture();
            document.AddParagraph()
                    .AppendLine("Lets add another picture (without the fancy  rotation stuff)")
                    .AppendLine()
                    .AppendPicture(pictureNormal);

            // Save this document.
            document.Save();
        }

        private static void HelloWorldAdvancedFormatting()
        {
            Enter();

            var document = Document.Create("HelloWorldAdvancedFormatting.docx");

            // Insert a new Paragraphs.
            Paragraph p = document.AddParagraph();

            p.Append("I am ").Append("bold").Bold()
            .Append(" and I am ")
            .Append("italic").Italic().Append(".")
            .AppendLine("I am ")
            .Append("Arial Black")
            .Font(new FontFamily("Arial Black"))
            .Append(" and I am not.")
            .AppendLine("I am ")
            .Append("BLUE").Color(Color.Blue)
            .Append(" and I am ")
            .Append("Red").Color(Color.Red).Append(".");

            // Save this document.
            document.Save();
        }

        private static void HelloWorldProtectedDocument()
        {
            Enter();

            var document = Document.Create("unused.docx");

            // Insert a Paragraph into this document.
            document.AddParagraph("Hello, World!")
                    .Font(new FontFamily("Times New Roman"))
                    .FontSize(32)
                    .Color(Color.Blue)
                    .Bold();

            // Save this document to disk with different options

            // Protected with password for Read Only
            document.AddProtection(EditRestrictions.ReadOnly, "password");
            document.SaveAs("HelloWorldPasswordProtectedReadOnly.docx");

            // Protected with password for Comments
            document.AddProtection(EditRestrictions.Comments, "password");
            document.SaveAs("HelloWorldPasswordProtectedCommentsOnly.docx");

            // Protected with password for Forms
            document.AddProtection(EditRestrictions.Forms, "password");
            document.SaveAs("HelloWorldPasswordProtectedFormsOnly.docx");

            // Protected with password for Tracked Changes
            document.AddProtection(EditRestrictions.TrackedChanges, "password");
            document.SaveAs("HelloWorldPasswordProtectedTrackedChangesOnly.docx");

            // Protected with password for Read Only
            document.AddProtection(EditRestrictions.ReadOnly);
            document.SaveAs("HelloWorldWithoutPasswordReadOnly.docx");

            // Protected with password for Comments
            document.AddProtection(EditRestrictions.Comments);
            document.SaveAs("HelloWorldWithoutPasswordCommentsOnly.docx");

            // Protected with password for Forms
            document.AddProtection(EditRestrictions.Forms);
            document.SaveAs("HelloWorldWithoutPasswordFormsOnly.docx");

            // Protected with password for Tracked Changes
            document.AddProtection(EditRestrictions.TrackedChanges);
            document.SaveAs("HelloWorldWithoutPasswordTrackedChangesOnly.docx");
        }

        private static void HighlightWords()
        {
            Enter();

            var document = Document.Create("HighlightWords.docx");

            document.AddParagraph("First line. ")
                .Append("This sentence is highlighted")
                    .Highlight(Highlight.Yellow)
                .Append(", but this is ")
                .Append("not").Italic()
                .Append(".");

            document.Save();
        }

        /// <summary>
        /// Creates a document with a Hyperlink, an Image and a Table.
        /// </summary>
        private static void HyperlinksInDocument()
        {
            Enter();

            var document = Document.Create("Hyperlinks.docx");

            // Add a title
            document.AddParagraph("Test")
                    .Align(Alignment.Center)
                    .FontSize(20)
                    .Font(new FontFamily("Comic Sans MS"));

            // Insert a new Paragraph into the document.
            document.AddParagraph()
                    .AppendLine("This line contains a ")
                    .Append("bold ").Bold().Append("word.")
                    .AppendLine()
                    .AppendLine("And a ")
                    .Append(new Hyperlink("link", new Uri("http://www.microsoft.com")))
                    .Append(".");

            // Insert a hyperlink into the paragraph
            string text = "A final paragraph - ";
            var p = document.AddParagraph(text)
                .AppendLine("With a few lines of text to read.")
                .AppendLine("And a final line.");

            p.InsertHyperlink(new Hyperlink("second link", new Uri("http://docs.microsoft.com/")), text.Length);

            // Save this document.
            document.Save();
        }

        /// <summary>
        /// Create a document with a Paragraph where the first line is indented.
        /// </summary>
        private static void Indentation()
        {
            Enter();

            var document = Document.Create("Indentation.docx");

            Paragraph p = document.AddParagraph("Line 1\nLine 2\nLine 3");
            p.IndentationFirstLine = 1.0f;
            document.Save();
        }

        private static void LargeTable()
        {
            TableBorder noTableBorder = new TableBorder(TableBorderStyle.None, 0, 0, Color.White);

            Enter();

            var doc = Document.Create("LargeTables.docx");
            Table table = doc.AddTable(1, 18);

            double wholeWidth = doc.PageWidth - doc.MarginLeft - doc.MarginRight;
            double colWidth = wholeWidth / table.ColumnCount;
            table.AutoFit = AutoFit.Contents;
            Row row = table.Rows[0];
            List<Cell> cells = row.Cells.ToList();

            for (int i = 0; i < cells.Count; i++)
            {
                cells[i].Paragraphs[0].Append($"Column #{i}");
                cells[i].Width = colWidth;
                cells[i].BottomMargin = 0;
                cells[i].LeftMargin = 0;
                cells[i].RightMargin = 0;
                cells[i].TopMargin = 0;
            }

            table.SetBorder(TableBorderType.Bottom, noTableBorder);
            table.SetBorder(TableBorderType.Left, noTableBorder);
            table.SetBorder(TableBorderType.Right, noTableBorder);
            table.SetBorder(TableBorderType.Top, noTableBorder);
            table.SetBorder(TableBorderType.InsideV, noTableBorder);
            table.SetBorder(TableBorderType.InsideH, noTableBorder);

            doc.Save();
        }

        private static void LineChart()
        {
            Enter();

            var document = Document.Create("LineChart.docx");

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
            document.AddParagraph("Diagram").FontSize(20);
            document.InsertChart(c);
            document.Save();
        }

        private static void PieChart()
        {
            Enter();

            var document = Document.Create("PieChart.docx");

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
            document.AddParagraph("Diagram").FontSize(20);
            document.InsertChart(c);
            document.Save();
        }

        /// <summary>
        /// Loads a document 'Input.docx' and writes the text 'Hello World' into the first imbedded Image.
        /// This code creates the file 'Output.docx'.
        /// </summary>
        private static void ProgrammaticallyManipulateImbeddedImage()
        {
            Enter();

            var document = Document.Load(Path.Combine("..", "Input.docx"));

            // Make sure this document has at least one Image.
            if (document.Images.Count > 0)
            {
                Bitmap b;
                DXPlus.Image img = document.Images[0];

                // Write "Hello World" into this Image.
                using (Stream stm = img.GetStream(FileMode.Open, FileAccess.Read))
                {
                    b = new Bitmap(stm);
                    Graphics g = Graphics.FromImage(b);
                    g.DrawString("Hello, World", new Font("Tahoma", 20), Brushes.Blue, new PointF(0, 0));
                }

                // Save this Bitmap back into the document using a Create\Write stream.
                using (Stream stm = img.GetStream(FileMode.Create, FileAccess.Write))
                {
                    b.Save(stm, ImageFormat.Png);
                }
            }
            else
            {
                Console.WriteLine("The provided document contains no Images.");
            }

            // Save this document as Output.docx.
            document.SaveAs("Output.docx");
        }

        /// <summary>
        /// Create a document that with RightToLeft text flow.
        /// </summary>
        private static void RightToLeft()
        {
            Enter();

            var document = Document.Create("RightToLeft.docx");

            Paragraph p = document.AddParagraph("Hello World.");

            // Make this Paragraph flow right to left. Default is left to right.
            p.Direction = Direction.RightToLeft;

            // You don't need to manually set the text direction foreach Paragraph, you can just call this function.
            document.SetDirection(Direction.RightToLeft);

            // Save all changes made to this document.
            document.Save();
        }

        private static void Setup(string testFolder)
        {
            if (Directory.Exists(testFolder))
            {
                Directory.Delete(testFolder, true);
            }

            Directory.CreateDirectory(testFolder);

            Directory.SetCurrentDirectory(testFolder);
        }
        private static void TablesDocument()
        {
            Enter();

            var document = Document.Create("Tables.docx");

            Table table = new Table(new[,] {{"1", "2"}, {"3", "4"}})
                {
                    Design = TableDesign.ColorfulGrid, Alignment = Alignment.Center
                };

            // Add it once
            document.AddTable(table);

            // Add it again
            document.AddParagraph()
                    .AppendLine("Here's another copy of the table!")
                    .AppendLine()
                    .AddTableAfterSelf(table)
                    .Rows[0].Cells[0].ReplaceText("1", "One");

            // Insert a new Paragraph into the document.
            document.AddParagraph()
                    .AppendLine()
                    .AppendLine("Adding another table...");

            Table table1 = new Table(2, 2) {Design = TableDesign.MediumGrid1Accent1, Alignment = Alignment.Center};

            table1.Rows[0].Cells[0].Paragraphs[0].Append("One");
            table1.Rows[0].Cells[1].Paragraphs[0].Append("Two");
            table1.Rows[1].Cells[0].Paragraphs[0].Append("Three");
            table1.Rows[1].Cells[1].Paragraphs[0].Append("Four");

            var p = document.AddParagraph()
                .AppendLine()
                .AppendLine("The table should be right above this paragraph!");

            p.InsertTableBeforeSelf(table1);
            p.AppendLine();

            document.Save();
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
