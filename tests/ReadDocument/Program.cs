using System;
using System.Linq;
using DXPlus;

var doc = Document.Create(@"C:\Users\mark\onedrive\desktop\t.docx");

var comment = doc.CreateComment("James Arrow", "Testing 1.2.3");

var p1 = doc.AddParagraph("This is the first paragraph.")
    .Append("With multiple runs.")
    .Append("Last run");
doc.AddParagraph("This is the second and final paragraph");

var r = p1.Runs.ElementAt(1);
p1.AttachComment(comment, r);

foreach (var c in doc.Comments)
{
    Console.WriteLine(c);
    foreach (var p in c.Paragraphs)
        Console.WriteLine(p.Text);
}


doc.Save();