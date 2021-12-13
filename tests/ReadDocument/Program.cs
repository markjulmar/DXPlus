using DXPlus;
using DXPlus.Tests;

var doc = Document.Create(@"C:\Users\mark\onedrive\desktop\tc.docx");

doc.AddParagraph(Helpers.GenerateLoremIpsum());

var nd = doc.NumberingStyles.Create(NumberingFormat.Numbered);

doc.AddParagraph("Item #1").ListStyle(nd)
   .AddParagraph("Sub-item #1").ListStyle(nd, level:1)
   .AddParagraph("Sub-item #2").ListStyle(nd, level: 1);
doc.AddParagraph("Item #2").ListStyle(nd)
   .AddParagraph("With another paragraph").ListStyle()
   .AddParagraph("And another ..").ListStyle()
   .AddParagraph("With a quote").Style("IntenseQuote");

doc.AddParagraph(Helpers.GenerateLoremIpsum());

nd = doc.NumberingStyles.Create(NumberingFormat.Numbered);

doc.AddParagraph("Item #1").ListStyle(nd)
    .AddParagraph("Sub-item #1").ListStyle(nd, level: 1)
    .AddParagraph("Sub-item #2").ListStyle(nd, level: 1);

doc.AddParagraph("And finally ending with a final paragraph.");
doc.Save();
