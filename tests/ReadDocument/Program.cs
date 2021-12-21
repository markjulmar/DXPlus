using System;
using System.IO;
using DXPlus;

var doc = Document.Create(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "tdx.docx"));

Table t = new Table(rows:1, columns:3);

t.AddRow().MergeCells(0, 2);
t.AddRow();
t.AddRow();
t.AddRow().MergeCells(0, 3);
t.AddRow();

t.MergeCellsInColumn(2, 2, 2);

doc.AddTable(t);

doc.Save();

Console.WriteLine("Created document.");