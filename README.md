# DXPlus - DocX parser and formatter library

This library is a fork of [DocX](http://docx.codeplex.com/) which has been heavily modified in a variety of ways, including adding support for later versions of Word. The library allows you to load or create Word documents and edit them with .NET Core code. No Office interop libraries are used and the source is fully managed code.

![Build and Publish DXPlus library](https://github.com/markjulmar/DXPlus/workflows/Build%20and%20Publish%20DXPlus%20library/badge.svg)

The library is available on NuGet:

```
Install-Package Julmar.DxPlus -Version 1.2.1
```

> **Note**:
>
> The original DocX code has been purchased by Xceed and is now maintained in [their open source GitHub repo](https://github.com/xceedsoftware/DocX) with options for support.

## Working with documents

Word documents are primarily composed of sections, paragraphs and tables. There is always at least one section in every document which contains the main body. Other sections can be added to change page-level characteristics such as orientation or margins. In addition, headers and footers are held in their own sections.

Within a section the document has paragraphs and tables. Paragraphs have properties, which control formatting and visual characteristics, and _runs_ of text or drawings (images, videos, etc.) which provide the content. A run also has properties which provide fine-tuning for colors or fonts or even override the paragraph-level formatting. Here's a basic structure:

```
Document
  |               +-- Headers --- Paragraph(s)
  |               |
  +--- Section ---+-- Properties
          |       |
          |       +-- Footers --- Paragraph(s)
          |
          +--- Paragraph                 
          +--- Paragraph    
          +--- Paragraph --- Properties
                   |
                   +--- Run
                   +--- Run
                   +--- Run --- Properties
                         |
                         +--- Text / Drawing
                         +--- Text / Drawing
                         +--- Text / Drawing
```

### Create a new document

The library is oriented around the `IDocument` interface. It provides the basis for working with a single document. The primary namespace is `DXPlus`, and there's a secondary namespace for all the charting capabilities (`DXPlus.Charts`). Here is a synthesis of all the capabilities at the document level - this combines the interface along with all extension methods.

Documents are created and opened with the static `Document` class. You can open or create documents with static methods on this class.

```csharp
public class Document
{
    public static IDocument Load(string filename);
    public static IDocument Load(Stream stream);
    public static IDocument Create(string filename = null);
    public static IDocument CreateTemplate(string filename = null);
}
```

It has `Create` and `Open` methods which return the `IDocument` interface.

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

### Saving and closing documents

The `IDocument` interface implements `IDisposable`. Disposing the document is the same as closing it - it will release all the open resources. If you make changes to the document you need to call `Save` or `SaveAs` to commit those changes to the file or stream.

If `Save` is used, then the document must have a filename associated with it or an exception will be thrown.

```csharp
IDocument document = Document.Open("test.docx")
 ....

document.Save();
document.Close();
```

```csharp
using var document = Document.Create(); // can also pass a filename into Create
 ...
document.SaveAs("test.docx");
```

### Fluent API

Most of the API is _fluent_ in nature so each method returns the called object. This allows you to 'string together' changes to a paragraph, block, or section. Many of these methods are extension methods in the `DXPlus` namespace added to the `Paragraph`, `Document`, `Table` and `Picture` classes. Here's an example usage which creates a paragraph with an embedded hyperlink.

```csharp
var paragraph = new Paragraph()
    .AddText("This line contains a ")
    .Add(new Hyperlink("link", new Uri("http://www.microsoft.com")))
    .AddText(".");

// Add the paragraph to the document
doc.Add(paragraph);
```

### Add paragraphs

The most common action is to add paragraphs of text. This is done with the `Add` method. You can add additional text to the paragraph with the `AddText` method. A second parameter allows you to specify formatting such as text color, font, bold, etc. You can terminate a line with a carriage return with the `Newline` method. Here's an example:

```csharp
var document = Document.Create();
var p = document.Add("Hello, World! This is the first paragraph.")
    .Newline()
    .AddText("This is a second line. ")
    .AddText("It includes some ")
    .AddText("large", new Formatting {Font = new FontFamily("Times New Roman"), FontSize = 32})
    .AddText(", blue", new Formatting {Color = Color.Blue})
    .AddText(", bold text.", new Formatting {Bold = true})
    .Newline()
    .AddText("And finally some normal text.");
```

The above code will create a single paragraph with colors and various fonts. You can add multiple paragraphs by chaining to other `Add` methods, or add an empty paragraph with the document `AddParagraph` method as shown below.

```csharp
document.AddParagraph()
    .AddText("This sets the text of the paragraph")
    .Add("This adds a second paragraph")
    .Add("And a third.");

// Can add multiple paragraphs with AddRange
document.AddRange(new[] { "This is a paragraph", "This is too."});

// Can also create paragraphs directly. And use new C# features to condense code
document.Add(new Paragraph("I am centered 20pt Comic Sans.", new() { Font = new FontFamily("Comic Sans MS"), FontSize = 20 }) 
        { Properties = new() { Alignment = Alignment.Center } });
```

You can set specific styles and highlight text

```csharp
document.Add("Highlighted text").Style(HeadingType.Heading2);
document.Add("First line. ")
    .AddText("This sentence is highlighted", new() { Highlight = Highlight.Yellow })
    .AddText(", but this is ")
    .AddText("not", new() { Italic = true })
    .AddText(".");
```

Or add line indents through the `WithProperties` method.

```csharp
document.Add(new Paragraph("This paragraph has the first sentence indented. "
              + "It shows how you can use the Intent property to control how paragraphs are lined up.")
        { Properties = new() { FirstLineIndent = 20 } })
    .Newline()
    .AddText("This line shouldn't be indented - instead, it should start over on the left side.");
```

### Page breaks

Add a page break with the `AddPageBreak` method.

```csharp
document.AddPageBreak();
```

### Finding and Replacing text

`IDocument` has methods to locate text by string or `Regex` and optionally replace it. These methods walk through all paragraphs across all sections.

```csharp
IEnumerable<(Paragraph owner, int position)> results = document.FindText("look for me", StringComparison.CurrentCulture);
 ...
IEnumerable<(Paragraph owner, int position)> results = document.FindPattern(new Regex("^The"));


bool foundText = document.FindReplace("original text", "replacement text", StringComparison.CurrentCulture);
 ...

// Can remove located text
document.FindReplace("original text", null, StringComparison.CurrentCulture);
```

`Paragraph` has similar methods which are scoped to that paragraph.

```csharp
IEnumerable<int position> results = paragraph.FindText("look for me", StringComparison.CurrentCulture);
 ...
IEnumerable<int position> results = paragraph.FindPattern(new Regex("^The"));
 ...
bool foundText = paragraph.FindReplace("original text", "replacement text", StringComparison.CurrentCulture);
```

### Hyperlinks

Hyperlinks can be added to paragraphs - this creates a clickable element which can point to an external source, or to a section of the document.

```csharp
var paragraph = document.Add("This line contains a ")
    .Add(new Hyperlink("hyperlink", new Uri("http://www.microsoft.com")))
    .AddText(". Here's a .");

// Insert another hyperlink into the paragraph.
paragraph.Insert(p.Text.Length - 2, new Hyperlink("second link", new Uri("http://docs.microsoft.com/")));
```

### Images

Images such as `.png`, `.jpg`, or `.tiff` can be inserted into the document. The binary image data must be added to the document first - it's stored as a blob which can then be inserted zero or more times into paragraphs. Each time you insert the image, it's wrapped in a `Picture` object. The picture has properties to control the size, shape, rotation, etc. which are all applied to the image data for that specific render. The picture, in turn, is held in a `Drawing` element which is what actually gets inserted into the `Run`. Here's the relationship structure:

```
Run
 |
 +--- Drawing
        |
        +--- Picture
                |
                +----> Image (.bmp, .jpg, etc.)
```

Inserting an image involves three steps:

1. Create an `Image` object that wraps an image file (.png, .jpeg, etc.)
1. Create a `Picture` object from the image to set attributes such as shape, size, and rotation.
1. Insert the picture into a paragraph. This will automatically wrap the picture in a `Drawing`.

```csharp
// Add an image into the document.
var image = document.CreateImage("images/comic.jpg");

// Create a picture
Picture picture = image.CreatePicture(189, 128)
    .SetPictureShape(BasicShapes.Ellipse)
    .SetRotation(20)
    .SetDecorative(true)
    .SetName("Bat-Man!");

// Insert the picture into the document
document.Add("Just below there should be a picture rotated 10 degrees.")
    .Add(picture)
    .Newline();

// Add a second copy of the same image by creating a new picture
// Here we pass a name and description but omit the width/height so the native image dimensions are used.
document.Add("Second copy without rotation.")
    .Add(image.CreatePicture("My Favorite Superhero", "This is a comic book"));

// You can also grab the owner Drawing to add a caption or set other properties.
Picture finalPicture = image.CreatePicture(string.Empty, string.Empty);
document.Paragraphs.Last().Add(finalPicture);

// Drawing must be in document to add a text caption under it.
finalPicture.Drawing.AddCaption("The batman!");
```

### Styles

Default and custom styles can be applied to text.

```csharp
// Set the style of the text
document.Add("Styled Text").Style(HeadingType.Heading2);

// Can also set through properties.
var paragraph = document.Add("This is the title");
paragraph.Properties.StyleName = HeadingType.Title;
```

### Headers and Footers

You can add headers or footers to the first page, even pages, or default (used as odd if even is supplied). This is done through _sections_. There's always at least one section in the document, and each section can have a different set of headers/footers.

The header/footer itself is a container so you can add paragraphs, images, etc. to it.

```csharp
var mainSection = document.Sections.First();

var footer = mainSection.Footers.Default;
footer.MainParagraph.Properties = new() {Alignment = Alignment.Right};
footer.MainParagraph.Text = "Page ";
footer.MainParagraph.AddPageNumber(PageNumberFormat.Normal);

var image = document.CreateImage(Path.Combine("images", "clock.png"));
var picture = image.CreatePicture(48, 48);
picture.IsDecorative = true;

var header = mainSection.Headers.Default;
header.MainParagraph.Text = "Welcome to the ";
header.MainParagraph.Add(picture);
header.MainParagraph.AddText(" tower!");
```

### Lists

The library supports both numbered and bulleted lists. In Word, lists are just paragraphs with a specific style applied that adds the number or bullet prefix to each item.

> **Note:** If the styles aren't present in the document, they are automatically added.

To add a list, start by creating a specific style with the `NumberingStyles.Create` method exposed on the document. This style is then added to each paragraph you want to include in the list.

#### Numbered lists

```csharp
document.AddPageBreak();

document.Add("Numbered List").Style(HeadingType.Heading2);

var numberStyle = document.NumberingStyles.Create(NumberingFormat.Numbered);
document.Add("First Item").ListStyle(numberStyle)
    .AddParagraph("First sub list item").ListStyle(numberStyle, level: 1)
    .AddParagraph("Second item.").ListStyle(numberStyle)
    .AddParagraph("Third item.").ListStyle(numberStyle)
    .AddParagraph("Nested item.").ListStyle(numberStyle, level: 1)
    .AddParagraph("Second nested item.").ListStyle(numberStyle, level: 1);
```

#### Bulleted lists

The same code will create a bulleted list if you specify a bulleted format:

```csharp
document.Add("Bullet List").Style(HeadingType.Heading2);

var numberStyle = document.NumberingStyles.Create(NumberingFormat.Bullet);
document.Add("First Item").ListStyle(numberStyle)
    .AddParagraph("First sub list item").ListStyle(numberStyle, level: 1)
    .AddParagraph("Second item.").ListStyle(numberStyle)
    .AddParagraph("Third item.").ListStyle(numberStyle)
    .AddParagraph("Nested item.").ListStyle(numberStyle, level: 1)
    .AddParagraph("Second nested item.").ListStyle(numberStyle, level: 1);
```

You can also modify the font characteristics of the list.

```csharp
const double fontSize = 15;
document.Add("Lists with fonts").Style(HeadingType.Heading2);
var style = document.NumberingStyles.Create(NumberingFormat.Bullet);

foreach (var fontFamily in FontFamily.Families.Take(20))
{
    document.Add(new Paragraph(fontFamily.Name, 
            new() {Font = fontFamily, FontSize = fontSize})
        .ListStyle(style));
}
```

### Tables

Tables can be inserted into paragraphs with the `InsertAfter` or `InsertBefore` methods. A single paragraph can only have one table following it - you can determine if there is a table after the paragraph through the `Table` property. Word does not allow two tables be be placed sequentially in a document - they are merged into a single table.

```csharp
document.Add("Basic Table").Style(HeadingType.Heading2)

// Construct a 2x5 table
Table table = new Table(new[,]
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

// Add the table
document.AddParagraph().InsertAfter(table);
```

You can set margins, borders or other properties of the table through fluent methods.

```csharp
doc.Add("Large 10x10 table across whole page width").Style(HeadingType.Heading2);

// Create a 10x10 table with empty cells.
Table table = new Table(10, 10);

// Determine the width of the page
var section = doc.Sections.First();
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
```

### Bookmarks

You can set bookmarks into the document and then reference them by name to modify parts of the text.

```csharp
// Add a bookmark
var paragraph = doc
    .Add("This is a paragraph which contains a ")
    .AddBookmark("namedBookmark")
    .AddText("bookmark");

// Set text at the bookmark
paragraph.InsertTextAtBookmark("namedBookmark", "handy ");

Console.WriteLine(paragraph.Text); // "This is a paragraph which contains a handy bookmark"
```

You can also insert a bookmark across multiple runs to get the text for those runs - even if they change over time.

```csharp
var paragraph = document
    .Add("This is a test paragraph.")
    .AddText(" With lots of text.")
    .AddText(" Added over time.")
    .AddText(" And a final sentence.");

var runs = paragraph.Runs.ToList();
paragraph.SetBookmark("bookmark1", runs[1], runs[2]);
var bookmark = paragraph.Bookmarks[0];

Console.WriteLine(bookmark.Text); // " With lots of text. Added over time."
```

### Equations

Math equations are a special style that uses a monospaced font and provides a built-in equation editor included in Word. These can be inserted at the document level where it creates a new paragraph, or at the end of a paragraph.

```csharp
document.AddEquation("x = y+z");

// Blue, large equation
var paragraph = new Paragraph();
paragraph.AddEquation("x = (y+z)/t", new Formatting {FontSize = 18, Color = Color.Blue});
document.Add(paragraph);
```

### Document properties

There are two types of document properties you can work with.

1. Standard properties
1. Custom properties

#### Standard properties

Standard document properties are the standard metadata included in a document. They are exposed through the `Properties` collection on the document itself.

```csharp
document.Properties.Creator = "John Smith";
document.Properties.Title = "Test document created by C#";
```

#### Custom properties

Custom properties allow you to define additional metadata in the document stored as various data types such as Integer, String, Date, etc. These are managed in the `CustomProperties` collection exposed on the document.

```csharp
document.CustomProperties.Add("ReplaceMe", " inserted field ");
  ...

if (document.CustomProperties.TryGetValue("ReplaceMe", out CustomProperty property))
{
    Console.WriteLine(property.Name); // "ReplaceMe"
    Console.WriteLine(property.Value); // " inserted field "
}

```

The `CustomProperties` collection exposes an `IList<CustomProperty>` - this can be manipulated to add or remove properties in addition to the simple actions exposed above. This includes an `As<T>` method to cast the value to an expected type based on the `Type` property.

```csharp
foreach (CustomProperty property in document.CustomProperties)
{
    Console.WriteLine(property.Name); // ReplaceMe, ...
    Console.WriteLine(property.Type); // CustomPropertyType.Text, CustomPropertyType.Integer, etc.
}

document.CustomProperties.Clear(); // empty

document.CustomProperties.Add(new CustomProperty("Total NetWorth", 1000.0));
document.CustomProperties.Add(new CustomProperty("LastScan", DateTime.Now));

var prop = document.CustomProperties[0]; // Total NetWorth
string text = prop.Value; // "1000"
double value = prop.As<double>(); // 1000.0
```

### Using replacement fields

Both core properties and custom properties support replacement fields in the document where the _value_ is stored in the metadata. 
Once defined, you can inject standard or custom properties into the document text.

```csharp
document.Add("This paragraph has an")
    .AddCustomPropertyField("ReplaceMe")
    .AddText("which was added by ")
    .AddDocumentPropertyField(DocumentPropertyName.Creator)
    .AppendLine(".");

Console.WriteLine(document.Text); // "This paragraph has an inserted field which was added by John Smith."
```

### Charts

The library includes support for a few of the chart types in Word.

1. BarChart
1. PieChart
1. LineChart


```csharp
var chart = new BarChart();
...
document.AddParagraph().Add("chart1", chart); // adds to new paragraph
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
	new CompanySales { Year = "2016", TotalSales = 1.2 },
	new CompanySales { Year = "2017", TotalSales = 2.4 },
	new CompanySales { Year = "2018", TotalSales = 3.6 },
	new CompanySales { Year = "2019", TotalSales = 5.8 },
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
document.AddParagraph().Add(chart);
```

#### Pie chart

```csharp
PieChart chart = new PieChart();
chart.AddLegend(ChartLegendPosition.Bottom, false);

Series acmeSeries = new Series("ACME");
acmeSeries.Bind(acmeInc, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));
chart.AddSeries(acmeSeries);

document.AddParagraph().Add(chart);
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

document.AddParagraph().Add(chart);
```

#### Chart3D

The 3D chart is a `BarChart` with the `View3D` property set.

```csharp
var acmeSeries = new Series("ACME") { Color = Color.DarkBlue };
acmeSeries.Bind(acmeInc, nameof(CompanySales.Year), nameof(CompanySales.TotalSales));

var barChart = new BarChart { View3D = true };
barChart.AddSeries(acmeSeries);

// Insert chart into document
document.AddParagraph().Add(chart);
```
