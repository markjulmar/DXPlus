using System;
using System.Drawing;
using System.Linq;
using DXPlus;

var doc = Document.Create(@"C:\Users\mark\onedrive\desktop\t.docx");

var s = doc.Styles.AddStyle("Code", StyleType.Paragraph);
s.ParagraphFormatting.ShadeFill = Color.FromArgb(0xf0, 0xf0, 0xf0);
s.ParagraphFormatting.SetBorders(BorderStyle.Single, Color.LightGray, 5);
s.Formatting.Font = new FontFamily("Consolas");

var p1 = doc.AddParagraph("This is the first paragraph.")
    .Append("With multiple runs.").Append("Last run").Style(s);

doc.AddParagraph("This is the second and final paragraph");

p1.Runs.Last().Properties.ShadePattern = ShadePattern.DiagonalCross;
p1.Runs.Last().Properties.ShadeColor = Color.Blue;

doc.Save();