# DXPlus - DocX parser and formatter library

This library is a fork of [DocX](http://docx.codeplex.com/) which has been heavily modified in a variety of ways, including adding support for later versions of Word. The library allows you to load or create Word documents and edit them with .NET Core code. No Office interop libraries are used and the source is fully managed code.

![Build and Publish DXPlus library](https://github.com/markjulmar/DXPlus/workflows/Build%20and%20Publish%20DXPlus%20library/badge.svg)

The library is available on NuGet as an unlisted entry.

```
Install-Package Julmar.DxPlus -Version 1.0.0-prerelease
```

> **Note**:
>
> The original DocX library has been purchased by Xceed and is now maintained in [their GitHub repo](https://github.com/xceedsoftware/DocX) with full support. If you want a fully supported library with the same features, definitely check out their version.

## Usage

The library is oriented around the `IDocument` interface. It provides the basis for working with a single document. The primary namespace is `DXPlus`, and there's a secondary namespace for all the charting capabilities (`DXPlus.Charts`).

```csharp
public interface IDocument : IDisposable
{
	// Document properties
    public string RevisionId { get; }
    public IReadOnlyDictionary<DocumentPropertyName, string> DocumentProperties { get; }
    public IReadOnlyDictionary<string, object> CustomProperties { get; }
    public void SetPropertyValue(DocumentPropertyName name, string value);
    public void AddCustomProperty(string name, string value);
    public void AddCustomProperty(string name, double value);
    public void AddCustomProperty(string name, bool value);
    public void AddCustomProperty(string name, DateTime value);
    public void AddCustomProperty(string name, int value);

    // Save/Close
    public void Close();
    public void Save();
    public void SaveAs(string newFileName);
    public void SaveAs(Stream newStreamDestination);
    public void Dispose();

    // Methods to work with paragraphs, sections, and pages
    public IReadOnlyList<Paragraph> Paragraphs { get; }
    public Paragraph InsertParagraph(int index, Paragraph paragraph);
    public Paragraph AddParagraph(Paragraph paragraph);
    public Paragraph InsertParagraph(int index, string text, Formatting formatting);
    public Paragraph AddParagraph(string text, Formatting formatting);
    public bool RemoveParagraph(int index);
    public bool RemoveParagraph(Paragraph paragraph);

    // Elements included in paragraphs
    public IEnumerable<Hyperlink> Hyperlinks { get; }
    public NumberingStyleManager NumberingStyles { get; }
    public StyleManager Styles { get; }

    // Sections
    public IReadOnlyList<Section> Sections { get; }
    public void AddSection();
    public void AddPageBreak();
    public bool DifferentEvenOddHeadersFooters { get; set; }
    public IEnumerable<string> EndnotesText { get; }
    public IEnumerable<string> FootnotesText { get; }

    // Global replace text
    public void ReplaceText(string searchValue, string newValue, RegexOptions options,
        Formatting newFormatting, Formatting matchFormatting, MatchFormattingOptions formattingOptions, bool escapeRegEx, bool useRegExSubstitutions);

    // Bookmarks
    public BookmarkCollection Bookmarks { get; }
    public bool InsertAtBookmark(string bookmarkName, string toInsert);

    // Tables
    public IEnumerable<Table> Tables { get; }
    public Table AddTable(Table table);
    public Table InsertTable(int index, Table table);
    
    // Lists
    public IEnumerable<List> Lists { get; }
    public List AddList(List list);
    public List InsertList(int index, List list);

    // Images
    public List<Image> Images { get; }
    public Image AddImage(string imageFileName);
    public Image AddImage(Stream imageStream, string contentType = "image/jpg");

    // Pictures (in the paragraph)
    public IEnumerable<Picture> Pictures { get; }

    // Document template support
    public void ApplyTemplate(string templateFilePath);
    public void ApplyTemplate(string templateFilePath, bool includeContent);
    public void ApplyTemplate(Stream templateStream);
    public void ApplyTemplate(Stream templateStream, bool includeContent);

    // Charts
    public void InsertChart(Chart chart);
    
    // Table of contents
    public TableOfContents InsertDefaultTableOfContents();
    public TableOfContents InsertTableOfContents(Paragraph reference, string title, 
    		TableOfContentsSwitches switches, string headerStyle, int maxIncludeLevel, int? rightTabPos);
}

```

### Create a new document

Documents are created and opened with the static `Document` class. You can open or create documents with this static class.

```csharp
public static class Document
{
    public static IDocument Load(string filename);
    public static IDocument Load(Stream stream);
    public static IDocument Create(string filename = null);
    public static IDocument CreateTemplate(string filename = null);
}
```

It has `Create` and `Open` methods which then return the `IDocument` interface.

```csharp
IDocument document = Document.Create("test.docx"); // named -- but not written until Save is called.

...

document = Document.Create(); // No name -- must use SaveAs
```

### Open a document

You can open a document from a file or a stream (including a network stream).

```csharp
// Can read from an existing local document.
IDocument document = Document.Open("test.docx")

...

// Can read from a stream.
var document = Document.Open(await client.ReadAsStreamAsync())
```

### Fluent API

Most of the API is _fluent_ in nature so each method returns the called object. This allows you to 'string together' changes to a paragraph, block, or section. These are all .NET extension methods in the `DXPlus` namespace added to the `Paragraph`, `Document`, `Table` and `Picture` classes.

### Saving changes

The document can be saved with the `Save` or `SaveAs` method. If `Save` is used, then the document cannot be new or an exception will be thrown. If you don't want to save the document, you can call `Close` or `Dispose` to throw away changes.

```csharp
IDocument document = Document.Open("test.docx")
 ....

document.Save();
```

### Add paragraphs

The most common action is to add paragraphs of text. This is done with the `AddParagraph` method. You can add text, lines, or other content to the paragraph with the `Append` methods.

```csharp
document.AddParagraph("Hello, World! This is the first paragraph.")
    .AppendLine()
    .Append("It includes some ")
    .Append("large").WithFormatting(new Formatting { Font = new FontFamily("Times New Roman"), FontSize = 32 })
    .Append(", blue").WithFormatting(new Formatting { Color = Color.Blue })
    .Append(", bold text.").WithFormatting(new Formatting { Bold = true })
    .AppendLine()
    .AppendLine("And finally some normal text.");
```

As shown above, you can change the text attributes (color, font, size, etc.) of the text using the `WithFormatting` method.

```csharp
document.AddParagraph()
    .Append("I am ")
    .Append("bold").WithFormatting(new Formatting { Bold = true })
    .Append(" and I am ")
    .Append("italic").WithFormatting(new Formatting { Italic = true })
    .Append(".").AppendLine()
    .AppendLine("I am ")
    .Append("Arial Black").WithFormatting(new Formatting { Font = new FontFamily("Arial Black") })
    .Append(" and I am ").AppendLine()
    .Append("Blue").WithFormatting(new Formatting { Color = Color.Blue })
    .Append(" and I am ")
    .Append("Red").WithFormatting(new Formatting { Color = Color.Red })
    .Append(".");

document.AddParagraph("I am centered 20pt Comic Sans.")
    .WithProperties(new ParagraphProperties {Alignment = Alignment.Center})
    .WithFormatting(new Formatting {Font = new FontFamily("Comic Sans MS"), FontSize = 20});
```

You can add a blank line with the same method

```csharp
// Blank line
document.AddParagraph();
```

You can also hilight text.

```csharp
document.AddParagraph("Highlighted text").Style(HeadingType.Heading2);
document.AddParagraph("First line. ")
    .Append("This sentence is highlighted").WithFormatting(new Formatting { Highlight = Highlight.Yellow })
    .Append(", but this is ")
    .Append("not").WithFormatting(new Formatting { Italic = true })
    .Append(".");
```

Or add line indents through the `WithProperties` method.

```csharp
document.AddParagraph("This paragraph has the first sentence indented. "
                      + "It shows how you can use the Intent property to control how paragraphs are lined up.")
    .WithProperties(new ParagraphProperties { FirstLineIndent = 20 })
    .AppendLine()
    .AppendLine("This line shouldn't be indented - instead, it should start over on the left side.");
```

### Page breaks

Add a page break with the `AddPageBreak` method.

```csharp
document.AddPageBreak();
```

### Replacing text

The `IDocument.ReplaceText` method can be used to find and replace regular expressions or text in all paragraphs of the document.

```csharp
document.ReplaceText("original text", "replacement text", RegexOptions.IgnoreCase);
```

The method has several optional parameters that allow you to change the formatting of the replacement.

```csharp
document.ReplaceText("original text", "replacement text", RegexOptions.None, 
                        new Formatting { Bold = true });
```

You can also match formatting and only replace what aligns to the passed formatting options.

```csharp
document.ReplaceText("original bold text", "replacement text", RegexOptions.None, 
                        null, new Formatting { Bold = true });
```

### Hyperlinks

Hyperlinks can be added to paragraphs - this creates a clickable element which can point to an external source, or to a section of the document.

```csharp
var paragraph = document.AddParagraph("This line contains a ")
    .Append(new Hyperlink("link", new Uri("http://www.microsoft.com")))
    .Append(". With a few lines of text to read.")
    .AppendLine(" And a final line with a .");

// Insert another hyperlink into the paragraph.
paragraph.InsertHyperlink(new Hyperlink("second link", new Uri("http://docs.microsoft.com/")), p.Text.Length - 2);
```

### Images

Images can be inserted into the document in a two-step fashion.

1. Create an `Image` object that wraps an image file (PNG, JPEG, etc.)
1. Create a `Picture` object from the image to set attributes such as shape, size, and rotation.
1. Insert the picture into a paragraph.

```csharp
// Add an image into the document.
var image = document.AddImage("images/comic.jpg");

// Create a picture
Picture picture = image.CreatePicture(189, 128)
    .SetPictureShape(BasicShapes.Ellipse)
    .SetRotation(20)
    .IsDecorative(true)
    .SetName("Bat-Man!");

// Insert the picture into the document
document.AddParagraph()
    .AppendLine("Just below there should be a picture rotated 10 degrees.")
    .Append(picture)
    .AppendLine();

// Add a second copy of the same image by creating a new picture
document.AddParagraph()
    .AppendLine("Second copy without rotation.")
    .Append(image.CreatePicture("My Favorite Superhero", "This is a comic book"));
```

### Styles

Default and custom styles can be applied to text.

```csharp
// Set the style of the text
document.AddParagraph("Styled Text").Style(HeadingType.Heading2);

// Can also set through properties.
var paragraph = document.AddParagraph("This is the title");
paragraph.Properties.StyleName = HeadingType.Title;
```

### Headers and Footers

You can add headers or footers to the first page, even pages, or default (used as odd if even is supplied). This is done through _sections_. There's always at least one section in the document, and each section can have a different set of headers/footers.

The header/footer itself is a container so you can add paragraphs, images, etc. to it.

```csharp
// Add some text into the first page header
var section = document.Sections.First();
section.Headers.First
    		.Add().Append("First page header")
          	.WithFormatting(new Formatting() {Bold = true});

// Create a picture and add it to the default header (all other pages)
var image = document.AddImage(Path.Combine("..", "images", "bulb.png"));
var picture = image.CreatePicture(15, 15);
picture.IsDecorative = true;
section.Headers.Default.Add().Append(picture);
```

### Lists

The library supports both numbered and bulleted lists. If the styles aren't present in the document, they are automatically added.

#### Numered lists

```csharp
document.AddParagraph("Numbered List").Style(HeadingType.Heading2);
List numberedList = new List(NumberingFormat.Numbered)
    .AddItem("First item.")
    .AddItem("First sub list item", level: 1)
    .AddItem("Second item.")
    .AddItem("Third item.")
    .AddItem("Nested item.", level: 1)
    .AddItem("Second nested item.", level: 1);
document.AddList(numberedList);
```

#### Bulleted lists

```csharp
document.AddParagraph("Bullet List").Style(HeadingType.Heading2);
List bulletedList = new List(NumberingFormat.Bulleted)
    .AddItem("First item.")
    .AddItem("Second item")
    .AddItem("Sub bullet item", level: 1)
    .AddItem("Second sub bullet item", level: 1)
    .AddItem("Third item");
document.AddList(bulletedList);
```

You can also modify the font characteristics of the list.

```csharp
document.AddParagraph("Lists with fonts").Style(HeadingType.Heading2);
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
```

### Tables

Tables can be inserted into paragraphs with the `AddTable` or `InsertTableBefore` methods.

```csharp
document.AddParagraph("Basic Table").Style(HeadingType.Heading2).AppendLine();

var table = new Table(new[,] {{"Title", "The wonderful world of Disney"}, {"# visitors", "200,000,000 per year."}})
{
    Design = TableDesign.ColorfulList,
    Alignment = Alignment.Center
};

document.AddParagraph()
        .AddTable(table);
```

You can set margins, borders or other properties of the table through fluent methods.

```csharp
document.AddParagraph("Large 10x10 table across whole page width")
		.Style(HeadingType.Heading2);

var section = document.Sections.First();
Table table = document.AddTable(10,10);

// Determine the width of the page
double pageWidth = section.Properties.PageWidth - section.Properties.LeftMargin - section.Properties.RightMargin;
double colWidth = pageWidth / table.ColumnCount;

// Add some random data into the table
foreach (var cell in table.Rows.SelectMany(row => row.Cells))
{
    cell.Paragraphs[0].SetText(new Random().Next().ToString());
    cell.Width = colWidth;
    cell.SetMargins(0);
}

// Auto fit the table and set a border
table.AutoFit = AutoFit.Contents;
TableBorder border = new TableBorder(TableBorderStyle.DoubleWave, 0, 0, Color.CornflowerBlue);
table.SetBorders(border);

// Insert the table into the document
document.AddParagraph("This line should be above the table.");
var paragraph = document.AddParagraph("... and this line below the table.");
p.InsertTableBefore(table);
```

### Bookmarks

You can set bookmarks into the document and then reference them by name to modify parts of the text.

```csharp
var paragraph = document.AddParagraph("This is a paragraph which contains a ")
    					.AppendBookmark("namedBookmark")
						.Append("bookmark");

// Set the text at the bookmark
p.InsertAtBookmark("namedBookmark", "handy ");
```

### Equations

Math equations are a special style that uses a monospaced font. These, unlike most other textual elements, are inserted at the _document_ level.

```csharp
document.AddEquation("x = y+z");

// Blue, large equation
document.AddEquation("x = (y+z)/t").WithFormatting(new Formatting {FontSize = 18, Color = Color.Blue});
```

### Document properties

There are two types of document properties you can work with.

1. Standard properties
1. Custom properties

#### Standard properties

Standard document properties are identified with the `DocumentPropertyName` items. These are part of the metadata associated with the document.

```csharp
document.SetPropertyValue(DocumentPropertyName.Creator, "John Smith");
document.SetPropertyValue(DocumentPropertyName.Title, "Test document created by C#");
```

#### Custom properties

Custom properties allow you to define replacement fields in the document where the _value_ is stored in the metadata.

```csharp
document.AddCustomProperty("ReplaceMe", " inserted field ");
```

Once defined, you can inject standard or custom properties into the document text.

```csharp
document.AddParagraph("This paragraph has an")
    .AddCustomPropertyField("ReplaceMe")
    .Append("which was added by ")
    .AddDocumentPropertyField(DocumentPropertyName.Creator)
    .AppendLine(".");
```

### Charts

The library includes support for most of the supported charting constructs in Word. 

1. BarChart
1. PieChart
1. LineChart
1. Chart3D

To add a chart to a document, you use the `InsertChart` method.

```csharp
var chart = new BarChart();
...

document.InsertChart(chart);
```

#### Supplying data to charts

Charts display data as a _series_. Each series has an X and Y axis which must be tied to a specific piece of data.

For example, if we have some sales data organized by year:

```csharp
public class CompanySales
{
	public string Year { get; set; }
	// Sales in millions of units
	public double TotalSales { get; set; }
}

...

CompanySales[] acmeInc = {
	new CompanySales { Year = 2016, TotalSales = 1.2 },
	new CompanySales { Year = 2017, TotalSales = 2.4 },
	new CompanySales { Year = 2018, TotalSales = 3.6 },
	new CompanySales { Year = 2019, TotalSales = 5.8 },
	....
};

CompanySales[] sprocketsInc = { ... }

```

We could then bind this to different chart styles using the names of the properties.

#### Bar chart

Bar charts show columns or rows.

```csharp
var chart = new BarChart {
    BarDirection = BarDirection.Column,
    BarGrouping = BarGrouping.Standard,
    GapWidth = 400
};

chart.AddLegend(ChartLegendPosition.Bottom, false);

// Add series
var acmeSeries = new Series("ACME") { Color = Color.DarkBlue };
acmeSeries.Bind(acmeInc, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
chart.AddSeries(acmeSeries);

var sprocketsSeries = new Series("sprocketsInc") { Color = Color.FromArgb(1, 0xff, 0, 0xff)};
sprocketsSeries.Bind(sprocketsInc, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
chart.AddSeries(sprocketsSeries);

// Insert chart into document
document.InsertChart(chart);
```

#### Pie chart

```csharp
PieChart chart = new PieChart();
chart.AddLegend(ChartLegendPosition.Bottom, false);

Series acmeSeries = new Series("ACME");
acmeSeries.Bind(acmeInc, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
chart.AddSeries(acmeSeries);

document.InsertChart(chart);
```

#### Line chart

```csharp
LineChart chart = new LineChart();
chart.AddLegend(ChartLegendPosition.Bottom, false);

var acmeSeries = new Series("ACME") { Color = Color.DarkBlue };
acmeSeries.Bind(acmeInc, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
chart.AddSeries(acmeSeries);

var sprocketsSeries = new Series("sprocketsInc") { Color = Color.FromArgb(1, 0xff, 0, 0xff)};
sprocketsSeries.Bind(sprocketsInc, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
chart.AddSeries(sprocketsSeries);

document.InsertChart(chart);
```

#### Chart3D

The 3D chart is a `BarChart` with the `View3D` property set.

```csharp
var acmeSeries = new Series("ACME") { Color = Color.DarkBlue };
acmeSeries.Bind(acmeInc, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));

var barChart = new BarChart { View3D = true };
barChart.AddSeries(acmeSeries);

// Insert chart into document
document.InsertChart(barChart);
```
