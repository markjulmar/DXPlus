using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DXPlus;
using DXPlus.Charts;

using Color = System.Drawing.Color;

namespace GenerateSampleDoc
{
    public static class Program
    {
        private static readonly List<Action<IDocument>> testers = new()
        {
            AddHeaderAndFooter,
            WriteTitle, 
            WriteFirstParagraph,
            AddCustomList,
            AddVideoToDoc,
            AddImageToDoc,
            AddPageBreak,
            CreateTableWithList,
            CreateNumberedList,
            CreateBulletedList,
            CreateBasicTable,
            AddTableToDocument,
            AddBookmarkToDocument,
            AddBarChartToDocument,
            AddLineChartToDocument,
            AddPieChartToDocument,
            CreateCustomTableStyle,
            DumpStyles,
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
                .AddText("large", new Formatting { Font = new FontValue("Times New Roman"), FontSize = 32 })
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
                    new() { Font = new FontValue("Comic Sans MS"), FontSize = 20 })
                { Properties = new() { Alignment = Alignment.Center } });

            doc.Add(new Paragraph("This paragraph has the first sentence indented. "
                          + "It shows how you can use the Indent property to control how paragraphs are lined up.")
                    { Properties = new() { FirstLineIndent = Uom.FromInches(0.5) } })
                .Newline()
                .AddText("This line shouldn't be indented - instead, it should start over on the left side.");
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
            doc.Add("Numbered List").Style(HeadingType.Heading2);

            var numberStyle = doc.NumberingStyles.AddNumberedDefinition(1);
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

            doc.Add("Lists with fonts").Style(HeadingType.Heading2);
            var style = doc.NumberingStyles.AddBulletDefinition();

            foreach (var fontFamily in FontValue.FontFamilies.Take(20))
            {
                doc.Add(new Paragraph(fontFamily.Name, 
                        new() {Font = fontFamily, FontSize = fontSize})
                    .ListStyle(style));
            }
        }

        static void AddCustomList(IDocument doc)
        {
            var nd = doc.NumberingStyles.AddCustomDefinition("", new FontValue("Wingdings"));
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
                cell.Properties.CellWidth = colWidth;
                cell.Properties.SetMargins(0);
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
                Properties = new() {
                    Design = TableDesign.ColorfulGridAccent6,
                    ConditionalFormatting = TableConditionalFormatting.FirstRow,
                    Alignment = Alignment.Center
                }
            };

            doc.AddParagraph().InsertAfter(table);
        }

        static void CreateTableWithList(IDocument doc)
        {
            doc.Add("This is a table.");

            var table = new Table(rows: 2, columns: 2);
            table.SetOutsideBorders(new Border(BorderStyle.Single, Uom.FromPoints(1))) //1pt
                 .SetInsideBorders(new Border(BorderStyle.Single, Uom.FromPoints(1.5))); // 1.5pt

            doc.Add(table);

            var nd = doc.NumberingStyles.AddNumberedDefinition(1);

            foreach (var row in table.Rows)
            {
                for (int col = 0; col < row.ColumnCount; col++)
                {
                    var cell = row.Cells[col];
                    cell.Properties.Shading = new() {Fill = col % 2 == 0 ? Color.Pink : Color.LightBlue};

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
            var chart = new PieChart();
            chart.AddLegend(ChartLegendPosition.Bottom, false);

            // Add series
            var acmeSeries = new Series("ACME");
            acmeSeries.Bind(CompanySales.Acme, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
            chart.AddSeries(acmeSeries);

            // Insert chart into document
            doc.AddParagraph().Add("sales3", chart);
        }

        static void CreateCustomTableStyle(IDocument doc)
        {
            doc.Add("Custom table style").Style(HeadingType.Heading2);

            var style = doc.Styles.Add("MyTableStyle", "My Table Style", StyleType.Table);

            var border = new Border(BorderStyle.Single, Uom.FromPoints(.5))
                {Color = new ColorValue(Color.FromArgb(0x9C, 0xC2, 0xE5), ThemeColor.Accent5, 153)};

            style.ParagraphFormatting = new() {LineSpacingAfter = 0};

            style.TableFormatting = new() { RowBands = 1 };
            style.TableFormatting.SetOutsideBorders(border);
            style.TableFormatting.SetInsideBorders(border);

            style.TableCellFormatting = new() {VerticalAlignment = VerticalAlignment.Center};

            border = new Border(BorderStyle.Single, Uom.FromPoints(.5))
                { Color = new ColorValue(Color.FromArgb(0x5B,0x9B,0xD5), ThemeColor.Accent5) };

            // First row style
            style.TableStyles.Add(new TableStyle(TableStyleType.FirstRow)
            {
                Formatting = new() {Bold = true, Color = new ColorValue(Color.White, ThemeColor.Background1)},
                TableCellFormatting = new()
                {
                    TopBorder = border, BottomBorder = border,
                    LeftBorder = border, RightBorder = border,
                    Shading = new()
                    {
                        Color = ColorValue.Auto,
                        Pattern = ShadePattern.Clear,
                        Fill = new ColorValue(Color.FromArgb(0x5B,0x9B,0xD5), ThemeColor.Accent5)
                    }
                }
            });

            style.TableStyles.Add(new TableStyle(TableStyleType.LastRow)
            {
                Formatting = new() { Bold = true },
                TableCellFormatting = new()
                {
                    TopBorder = new Border(BorderStyle.Double, Uom.FromPoints(.5))
                        { Color = new ColorValue(Color.FromArgb(0x5B, 0x9B, 0xD5), ThemeColor.Accent5) }
                }
            });

            style.TableStyles.Add(new TableStyle(TableStyleType.FirstColumn) {
                Formatting = new() { Bold = true },
            });

            style.TableStyles.Add(new TableStyle(TableStyleType.LastColumn) {
                Formatting = new() { Bold = true },
            });

            var cellBand = new TableCellProperties
            {
                Shading = new()
                {
                    Color = ColorValue.Auto, 
                    Pattern = ShadePattern.Clear,
                    Fill = new ColorValue(Color.FromArgb(0xDE, 0xEA, 0xF6), ThemeColor.Accent5, 51)
                }
            };
            style.TableStyles.Add(new TableStyle(TableStyleType.BandedEvenRows) {TableCellFormatting = cellBand});

            var table = new Table(4, 2)
            {
                Properties =
                {
                    Design = "MyTableStyle",
                    ConditionalFormatting = TableConditionalFormatting.FirstRow | TableConditionalFormatting.FirstColumn
                }
            };

            table.Rows[0].Cells[0].Text = "Header 1";
            table.Rows[0].Cells[1].Text = "Header 2";
            table.Rows[1].Cells[0].Text = "First cell";
            table.Rows[1].Cells[1].Text = "Right cell";
            table.Rows[2].Cells[0].Text = "Left cell";
            table.Rows[2].Cells[1].Text = "Last cell";
            table.Rows[3].Cells[0].Text = "Summary row";

            table.Rows[3].MergeCells(0,2);

            doc.Add(table);
        }

        static void AddBookmarkToDocument(IDocument doc)
        {
            doc.Add("Bookmark test").Style(HeadingType.Heading2);

            // Add a bookmark
            var paragraph = doc
                .Add("This is a paragraph which contains a ")
                .AddBookmark("namedBookmark")
                .AddText(" bookmark");

            // Set text at the bookmark
            paragraph.InsertTextAtBookmark("namedBookmark", "handy ");
        }

        static void AddPageBreak(IDocument doc)
        {
            doc.AddPageBreak();
            doc.Add("This page was intentionally blank")
                .Properties = new() {Alignment = Alignment.Center, LineSpacingBefore = Uom.FromPoints(288) };
            doc.AddPageBreak();
        }

        static void DumpStyles(IDocument doc)
        {
            doc.Add("Available Styles").Style(HeadingType.Heading2);

            var bulletList = doc.NumberingStyles.AddBulletDefinition();
            foreach (var style in doc.Styles)
            {
                doc.Add($"Id: {style.Id}, Name: {style.Name}, Type: {style.Type}, IsCustom: {style.IsCustom}")
                    .ListStyle(bulletList);
            }
        }
    }
}
