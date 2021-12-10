using DXPlus;

var doc = Document.Create(@"C:\Users\mark\onedrive\desktop\tc.docx");

doc.AddParagraph("This is a normal paragraph.");

var list = new List(NumberingFormat.Numbered);
list.AddItem("This is a list");
list.Items[0].Paragraph
    .AddParagraph("With another paragraph")
    .AddParagraph("And another ..")
    .AddParagraph("With a quote").Style("IntenseQuote");
list.AddItem("Starting the list again.")
    .AddItem("With another item (#3)");

doc.AddList(list);

doc.AddParagraph("And finally ending with a final paragraph.");

doc.Save();
