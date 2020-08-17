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
            //HighlightWords();
            //HelloWorldAdvancedFormatting();
            //HelloWorldProtectedDocument();
            //HelloWorldAddPictureToWord();
            //RightToLeft();
            //Indentation();
            //HeadersAndFooters();
            //HyperlinksInDocument();
            //AddList();
            //Equations();
            //Bookmarks();
            //BookmarksReplaceTextOfBookmarkKeepingFormat();
            //BarChart();
            //PieChart();
            //LineChart();
            //Chart3D();
            //DocumentMargins();
            //CreateTableWithTextDirection();
            //AddToc();
            //AddTocByReference();
            //TablesDocument();
            //DocumentsWithListsFontChange();
            //DocumentHeading();
            //LargeTable();
            //ProgrammaticallyManipulateImbeddedImage();
            //CountNumberOfParagraphs();
        }

        private static void CountNumberOfParagraphs()
        {
            Enter();

            DocX doc = DocX.Load(Path.Combine("..", "Input.docx"));

            foreach (Paragraph p in doc.Paragraphs)
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

            DocX document = DocX.Create("DocumentHeading.docx");

            foreach (HeadingType heading in (HeadingType[])Enum.GetValues(typeof(HeadingType)))
            {
                document.InsertParagraph($"{heading} - The quick brown fox jumps over the lazy dog")
                        .Heading(heading);
            }

            document.Save();
        }

        private static void AddList()
        {
            Enter();

            DocX document = DocX.Create("Lists.docx");
            List numberedList = document.CreateList()
                                        .AddItem("First List Item.", 0, ListItemType.Numbered, 2)
                                        .AddItem("First sub list item", 1)
                                        .AddItem("Second List Item.")
                                        .AddItem("Third list item.")
                                        .AddItem("Nested item.", 1)
                                        .AddItem("Second nested item.", 1);

            List bulletedList = document.CreateList()
                                        .AddItem("First Bulleted Item.", 0, ListItemType.Bulleted)
                                        .AddItem("Second bullet item")
                                        .AddItem("Sub bullet item", 1)
                                        .AddItem("Second sub bullet item", 1)
                                        .AddItem("Third bullet item");

            document.InsertList(numberedList);
            document.InsertList(bulletedList);
            document.Save();
        }

        private static void AddToc()
        {
            Enter();

            DocX document = DocX.Create("Toc.docx");

            document.InsertTableOfContents("I can haz table of contentz",
                TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2");

            document.InsertParagraph("Heading 1").Style("Heading1");

            document.InsertParagraph("Some very interesting content here");

            document.InsertParagraph("Heading 2").Style("Heading2");

            document.InsertSectionPageBreak();
            document.InsertParagraph("Some very interesting content here as well");

            document.InsertParagraph("Heading 3").Style("Heading3");

            document.InsertParagraph("Not so very interesting....");

            document.Save();
        }

        private static void AddTocByReference()
        {
            Enter();

            DocX document = DocX.Create("TocByReference.docx");

            document.InsertParagraph("Heading 1").Style("Heading1");
            document.InsertParagraph("Some very interesting content here");
            document.InsertParagraph("Heading 2").Style("Heading1");
            document.InsertSectionPageBreak();
            document.InsertParagraph("Some very interesting content here as well");
            Paragraph h2 = document.InsertParagraph("Heading 2.1").Style("Heading2");
            document.InsertParagraph("Not so very interesting....");

            document.InsertTableOfContents(h2, "I can haz table of contentz",
                TableOfContentsSwitches.O | TableOfContentsSwitches.U | TableOfContentsSwitches.Z | TableOfContentsSwitches.H, "Heading2");

            document.Save();
        }

        private static void BarChart()
        {
            Enter();

            DocX document = DocX.Create("BarChart.docx");

            // Create chart.
            BarChart chart = new BarChart
            {
                BarDirection = BarDirection.Column,
                BarGrouping = BarGrouping.Standard,
                GapWidth = 400
            };

            chart.AddLegend(ChartLegendPosition.Bottom, false);

            // Create data.
            List<ChartData> company1 = ChartData.CreateCompanyList1();
            List<ChartData> company2 = ChartData.CreateCompanyList2();

            // Create and add series
            Series series1 = new Series("Microsoft") { Color = Color.DarkBlue };
            series1.Bind(company1, nameof(ChartData.Month), nameof(ChartData.Money));
            chart.AddSeries(series1);

            Series series2 = new Series("Apple") { Color = Color.FromArgb(1, 0xff, 0, 0xff)};
            series2.Bind(company2, nameof(ChartData.Month), nameof(ChartData.Money));
            chart.AddSeries(series2);

            // Insert chart into document
            document.InsertParagraph("Diagram")
                    .FontSize(20);
            document.InsertChart(chart);

            document.Save();
        }

        private static void Bookmarks()
        {
            Enter();

            DocX document = DocX.Create("Bookmarks.docx");

            document.InsertBookmark("firstBookmark");

            Paragraph paragraph2 = document.InsertParagraph("This is a paragraph which contains a ")
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

            DocX docX = DocX.Load(Path.Combine("..", "DocumentWithBookmarks.docx"));

            foreach (Bookmark bookmark in docX.Bookmarks)
            {
                Console.WriteLine("Found bookmark {0}", bookmark.Name);
            }

            // Replace bookmars content
            docX.Bookmarks["bmkNoContent"].SetText("Here there was a bookmark");
            docX.Bookmarks["bmkContent"].SetText("Here there was a bookmark with a previous content");
            docX.Bookmarks["bmkFormattedContent"].SetText("Here there was a formatted bookmark");

            docX.SaveAs("BookmarksReplaceTextOfBookmarkKeepingFormat.docx");
        }

        private static void Chart3D()
        {
            Enter();

            DocX document = DocX.Create("3DChart.docx");

            BarChart c = new BarChart
            {
                View3D = true
            };

            // Create data.
            List<ChartData> company1 = ChartData.CreateCompanyList1();

            // Create and add series
            Series s = new Series("Microsoft")
            {
                Color = Color.GreenYellow
            };

            s.Bind(company1, nameof(ChartData.Month), nameof(ChartData.Money));
            c.AddSeries(s);

            // Insert chart into document
            document.InsertParagraph("3D Diagram").FontSize(20);
            document.InsertChart(c);

            document.Save();
        }

        private static void CreateTableWithTextDirection()
        {
            Enter();

            DocX document = DocX.Create("CeateTableWithTextDirection.docx");

            Table t = document.CreateTable(2, 3);
            t.Alignment = Alignment.Left;
            t.Design = TableDesign.MediumGrid1Accent2;

            List<Cell> cells = t.Rows[0].Cells.ToList();

            cells[0].Paragraphs[0].Append("A");
            cells[0].TextDirection = TextDirection.BottomToTopLeftToEnd;
            cells[1].Paragraphs[0].Append("B");
            cells[1].TextDirection = TextDirection.BottomToTopLeftToEnd;
            cells[2].Paragraphs[0].Append("C");
            cells[2].TextDirection = TextDirection.BottomToTopLeftToEnd;

            cells = t.Rows[1].Cells.ToList();
            cells[0].Paragraphs[0].Append("D");
            cells[1].Paragraphs[0].Append("E");
            cells[2].Paragraphs[0].Append("F");

            document.InsertTable(t);
            document.Save();
        }

        private static void DocumentMargins()
        {
            Enter();

            DocX document = DocX.Create("DocumentMargins.docx");

            // Create a float var that contains doc Margins properties.
            float leftMargin = document.MarginLeft;
            float rightMargin = document.MarginRight;
            float topMargin = document.MarginTop;
            float bottomMargin = document.MarginBottom;

            Console.WriteLine($"Current margins: L={leftMargin}, R={rightMargin}, T={topMargin}, B={bottomMargin}");

            // Modify using your own vars.
            leftMargin = 95F;
            rightMargin = 45F;
            topMargin = 50F;
            bottomMargin = 180F;

            // Or simply work the margins by setting the property directly. 
            document.MarginLeft = leftMargin;
            document.MarginRight = rightMargin;
            document.MarginTop = topMargin;
            document.MarginBottom = bottomMargin;

            // created bulleted lists

            List bulletedList = document.CreateList()
                .AddItem("First Bulleted Item.", 0, ListItemType.Bulleted)
                .AddItem("Second bullet item")
                .AddItem("Sub bullet item", 1)
                .AddItem("Second sub bullet item", 1)
                .AddItem("Third bullet item");


            document.InsertList(bulletedList);

            // Save this document.
            document.Save();
        }

        private static void DocumentsWithListsFontChange()
        {
            Enter();

            DocX document = DocX.Create("DocumentsWithListsFontChange.docx");

            foreach (FontFamily oneFontFamily in FontFamily.Families)
            {
                FontFamily fontFamily = oneFontFamily;
                const double fontSize = 15;

                // created numbered lists 
                List numberedList = document.CreateList()
                    .AddItem("First List Item.", 0, ListItemType.Numbered, 1)
                    .AddItem("First sub list item", 1)
                    .AddItem("Second List Item.")
                    .AddItem("Third list item.")
                    .AddItem("Nested item.", 1)
                    .AddItem("Second nested item.", 1);

                // created bulleted lists
                List bulletedList = document.CreateList()
                    .AddItem("First Bulleted Item.", 0, ListItemType.Bulleted)
                    .AddItem("Second bullet item")
                    .AddItem("Sub bullet item", 1)
                    .AddItem("Second sub bullet item", 1)
                    .AddItem("Third bullet item");

                document.InsertList(bulletedList);
                document.InsertList(numberedList, fontFamily, fontSize);
            }
            document.Save();
        }

        /// <summary>
        /// Create a document with two equations.
        /// </summary>
        private static void Equations()
        {
            Enter();

            DocX document = DocX.Create("Equations.docx");

            document.InsertEquation("x = y+z");

            document.InsertEquation("x = (y+z)/t")
                    .FontSize(18)
                    .Color(Color.Blue);

            document.Save();
        }

        private static void HeadersAndFooters()
        {
            Enter();

            DocX document = DocX.Create("HeadersAndFooters.docx");

            // Add Headers and Footers to this document.
            document.AddHeaders();
            document.AddFooters();

            // Force the first page to have a different Header and Footer.
            document.DifferentFirstPage = true;

            // Force odd & even pages to have different Headers and Footers.
            document.DifferentOddAndEvenPages = true;

            // Get the first, odd and even Headers for this document.
            Header headerFirst = document.Headers.First;
            Header headerOdd = document.Headers.Odd;
            Header headerEven = document.Headers.Even;

            // Get the first, odd and even Footer for this document.
            Footer footerFirst = document.Footers.First;
            Footer footerOdd = document.Footers.Odd;
            Footer footerEven = document.Footers.Even;

            // Insert a Paragraph into the first Header.
            headerFirst.InsertParagraph().Append("Hello First Header.").Bold();

            // Insert a Paragraph into the odd Header.
            headerOdd.InsertParagraph().Append("Hello Odd Header.").Bold();

            // Insert a Paragraph into the even Header.
            headerEven.InsertParagraph().Append("Hello Even Header.").Bold();

            // Insert a Paragraph into the first Footer.
            footerFirst.InsertParagraph().Append("Hello First Footer.").Bold();

            // Insert a Paragraph into the odd Footer.
            footerOdd.InsertParagraph().Append("Hello Odd Footer.").Bold();

            // Insert a Paragraph into the even Header.
            footerEven.InsertParagraph().Append("Hello Even Footer.").Bold();

            // Insert a Paragraph into the document.
            // Create a second page to show that the first page has its own header and footer.
            document.InsertParagraph().AppendLine("Hello First page.").InsertPageBreakAfterSelf();

            // Insert a Paragraph after the page break.
            Paragraph p7 = document.InsertParagraph().AppendLine("Hello Second page.");

            // Create a third page to show that even and odd pages have different headers and footers.
            p7.InsertPageBreakAfterSelf();

            // Insert a Paragraph after the page break.
            Paragraph p8 = document.InsertParagraph();
            p8.AppendLine("Hello Third page.");

            //Insert a next page break, which is a section break combined with a page break
            document.InsertSectionPageBreak();

            //Insert a paragraph after the "Next" page break
            Paragraph p9 = document.InsertParagraph();
            p9.Append("Next page section break.");

            //Insert a continuous section break
            document.InsertSection();

            //Create a paragraph in the new section
            Paragraph p10 = document.InsertParagraph();
            p10.Append("Continuous section paragraph.");

            // Save all changes to this document.
            document.Save();
        }

        /// <summary>
        /// Creates a simple document with the text Hello World.
        /// </summary>
        private static void HelloWorld()
        {
            Enter();

            DocX document = DocX.Create("HelloWorld.docx");
            Paragraph p = document.InsertParagraph();

            // Append some text and add formatting.
            p.Append("Hello World!")
                .Font(new FontFamily("Times New Roman"))
                .FontSize(32)
                .Color(Color.Blue)
                .Bold();

            p.AppendLine();
            p.AppendLine("This is some normal text.");

            document.Save();
        }

        /// <summary>
        /// Create a document with two pictures. One picture is inserted normal way, the other one with rotation
        /// </summary>
        private static void HelloWorldAddPictureToWord()
        {
            Enter();

            DocX document = DocX.Create("HelloWorldAddPictureToWord.docx");

            // Add an image into the document.    
            DXPlus.Image image = document.AddImage(Path.Combine("..", "images", "logo_template.png"));

            // Create a picture (A custom view of an Image).
            Picture picture = image.CreatePicture()
                                   .SetPictureShape(BasicShapes.Ellipse);
            picture.Rotation = 10;

            // Insert a new Paragraph into the document.
            Paragraph title = document.InsertParagraph()
                .Append("This is a test for a picture")
                .FontSize(20)
                .Font(new FontFamily("Comic Sans MS"));
            title.Alignment = Alignment.Center;

            // Insert a new Paragraph into the document.
            Paragraph p1 = document.InsertParagraph();

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
            document.InsertParagraph()
                    .AppendLine("Is it correct?")
                    .AppendLine();

            // Lets add another copy of the image
            Picture pictureNormal = image.CreatePicture();
            document.InsertParagraph()
                    .AppendLine("Lets add another picture (without the fancy  rotation stuff)")
                    .AppendLine()
                    .AppendPicture(pictureNormal);

            // Save this document.
            document.Save();
        }

        private static void HelloWorldAdvancedFormatting()
        {
            Enter();

            DocX document = DocX.Create("HelloWorldAdvancedFormatting.docx");

            // Insert a new Paragraphs.
            Paragraph p = document.InsertParagraph();

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

            DocX document = DocX.Create("unused.docx");

            // Insert a Paragraph into this document.
            document.InsertParagraph("Hello, World!")
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

            DocX document = DocX.Create("HighlightWords.docx");

            document.InsertParagraph("First line. ")
                .Append("This sentence is highlighted").Highlight(Highlight.Yellow)
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

            DocX document = DocX.Create("Hyperlinks.docx");

            // Add a title
            document.InsertParagraph("Test")
                    .Align(Alignment.Center)
                    .FontSize(20)
                    .Font(new FontFamily("Comic Sans MS"));

            // Insert a new Paragraph into the document.
            document.InsertParagraph()
                    .AppendLine("This line contains a ")
                    .Append("bold ").Bold().Append("word.")
                    .AppendLine()
                    .AppendLine("And this line has a cool ")
                    .AppendHyperlink(document.CreateHyperlink("link", new Uri("http://www.microsoft.com")))
                    .Append(".");

            // Save this document.
            document.Save();
        }

        /// <summary>
        /// Create a document with a Paragraph whos first line is indented.
        /// </summary>
        private static void Indentation()
        {
            Enter();

            DocX document = DocX.Create("Indentation.docx");

            Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
            p.IndentationFirstLine = 1.0f;
            document.Save();
        }

        private static void LargeTable()
        {
            Border noBorder = new Border(BorderStyle.None, 0, 0, Color.White);

            Enter();

            DocX doc = DocX.Create("LargeTables.docx");
            Table table = doc.InsertTable(1, 18);

            float wholeWidth = doc.PageWidth - doc.MarginLeft - doc.MarginRight;
            float colWidth = wholeWidth / table.ColumnCount;
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

            table.SetBorder(TableBorderType.Bottom, noBorder);
            table.SetBorder(TableBorderType.Left, noBorder);
            table.SetBorder(TableBorderType.Right, noBorder);
            table.SetBorder(TableBorderType.Top, noBorder);
            table.SetBorder(TableBorderType.InsideV, noBorder);
            table.SetBorder(TableBorderType.InsideH, noBorder);

            doc.Save();
        }

        private static void LineChart()
        {
            Enter();

            DocX document = DocX.Create("LineChart.docx");

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
            document.InsertParagraph("Diagram").FontSize(20);
            document.InsertChart(c);
            document.Save();
        }

        private static void PieChart()
        {
            Enter();

            DocX document = DocX.Create("PieChart.docx");

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
            document.InsertParagraph("Diagram").FontSize(20);
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

            DocX document = DocX.Load(Path.Combine("..", "Input.docx"));

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

            DocX document = DocX.Create("RightToLeft.docx");

            Paragraph p = document.InsertParagraph("Hello World.");

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

            DocX document = DocX.Create("Tables.docx");

            Table table = document.CreateTable(2, 2);

            table.Design = TableDesign.ColorfulGrid;
            table.Alignment = Alignment.Center;

            List<Cell> cells = table.Rows[0].Cells.ToList();
            cells[0].Paragraphs[0].Append("1");
            cells[1].Paragraphs[0].Append("2");

            cells = table.Rows[1].Cells.ToList();
            cells[0].Paragraphs[0].Append("3");
            cells[1].Paragraphs[0].Append("4");

            document.InsertParagraph()
                    .AppendLine("Can you check this Table of figures for me?")
                    .AppendLine()
                    .InsertTableAfterSelf(table);

            // Insert a new Paragraph into the document.
            document.InsertParagraph()
                    .AppendLine()
                    .AppendLine("Adding another table...");

            Table table1 = document.CreateTable(2, 2);

            table1.Design = TableDesign.ColorfulGridAccent2;
            table1.Alignment = Alignment.Center;

            table1.Rows[0].Cells.ElementAt(0).Paragraphs[0].Append("1");
            table1.Rows[0].Cells.ElementAt(1).Paragraphs[0].Append("2");
            table1.Rows[1].Cells.ElementAt(0).Paragraphs[0].Append("3");
            table1.Rows[1].Cells.ElementAt(1).Paragraphs[0].Append("4");

            Paragraph p = document.InsertParagraph()
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
