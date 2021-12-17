using System;
using System.Drawing;
using System.Linq;
using DXPlus;

var doc = Document.Create(@"C:\Users\mark\onedrive\desktop\tdx.docx");

doc.AddParagraph("Heading").Style(HeadingType.Heading1)
    .AddParagraph("This is a test");

Table t = new Table(3, 2)
{
    Design = TableDesign.LightListAccent1,
    TableCaption = "Welcome to the table",
    ConditionalFormatting = TableConditionalFormatting.FirstRow
};

// Header row
var boldFont = new Formatting {Bold = true};
t.Rows[0].Cells[0].Paragraphs.Single().SetText("Id", boldFont);
t.Rows[0].Cells[1].Paragraphs.Single().SetText("Value", boldFont);

for (int i = 0; i < 2; i++)
{
    t.Rows[i+1].Cells[0].Paragraphs.Single().SetText(100+i.ToString());
    t.Rows[i + 1].Cells[1].Paragraphs.Single().SetText($"Value #{i+1}");
}

doc.AddTable(t);
doc.AddParagraph("Final paragraph");

doc.Save();