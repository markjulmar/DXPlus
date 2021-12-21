using System;
using System.IO;
using DXPlus;

var doc = Document.Create(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "tdx.docx"));

doc.AddParagraph("First page!").Style(HeadingType.Heading1);
doc.AddParagraph("This is the first page of the document.");
doc.AddParagraph();

doc.AddSection(); // all paragraphs above this section.

doc.AddParagraph("This is a new section in landscape.");
doc.AddSection().Properties.Orientation = Orientation.Landscape;

doc.AddParagraph("Third page in portrait.");
doc.AddSection().Properties.Orientation = Orientation.Portrait;

doc.AddParagraph("Last page in portrait.");

doc.Save();

Console.WriteLine("Created document.");