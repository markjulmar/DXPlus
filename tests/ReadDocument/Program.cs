using System;
using System.IO;
using DXPlus;

string fn = @"C:\Users\mark\OneDrive\Desktop\introduction-winautomation.docx";

var doc = Document.Load(fn);

string title = string.Empty, author = string.Empty, summary = string.Empty;

foreach (var item in doc.Paragraphs)
{
    var styleName = item.Properties.StyleName;
    if (styleName == "Heading1") break;
    switch (styleName)
    {
        case "Title": title = item.Text; break;
        case "Author": author = item.Text; break;
        case "Abstract": summary = item.Text; break;
    }
}

Console.WriteLine($"Title: {title}");
Console.WriteLine($"Author: {author}");
Console.WriteLine($"Description: {summary}");

if (doc.DocumentProperties.TryGetValue(DocumentPropertyName.Creator, out var dtText))
{
    Console.WriteLine(dtText);
}
