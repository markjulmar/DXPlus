using System;
using System.IO;
using System.Linq;
using DXPlus;

public static class Program
{
    public static void Main(string[] args)
    {
        var doc = Document.Load(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "doc.docx"));

        //Console.WriteLine(doc.RawDocument());
        //Console.WriteLine();

        foreach (var p in doc.Paragraphs)
        {
            DumpParagraph(p, 0);
        }
    }

    private static void DumpParagraph(Paragraph block, int level)
    {
        string prefix = new string(' ', level*3);

        string listInfo = "";
        if (block.IsListItem())
        {
            listInfo = $"{block.GetNumberingFormat()} {block.GetListLevel()} #{block.GetListIndex()+1}";
        }

        Console.WriteLine($"{prefix}p: {block.Id} {block.Properties.StyleName} {listInfo}{block.Properties.DefaultFormatting}");
        foreach (var run in block.Runs)
        {
            DumpRun(run, level+1);
        }

        if (block.Table != null)
        {
            DumpTable(block.Table, level + 1);
        }
    }

    private static void DumpTable(Table table, int level)
    {
        string prefix = new string(' ', level*3);
        Console.WriteLine($"{prefix}tbl");
        foreach (var row in table.Rows)
        {
            DumpRow(row, level + 1);
        }
    }

    private static void DumpRow(Row row, int level)
    {
        string prefix = new string(' ', level * 3);
        Console.WriteLine($"{prefix}tr");
        foreach (var cell in row.Cells)
        {
            DumpCell(cell, level + 1);
        }
    }

    private static void DumpCell(Cell cell, int level)
    {
        string prefix = new string(' ', level * 3);
        Console.WriteLine($"{prefix}tc");
        foreach (var p in cell.Paragraphs)
        {
            DumpParagraph(p, level + 1);
        }
    }

    private static void DumpRun(Run run, int level)
    {
        string prefix = new string(' ', level*3);
        Console.WriteLine($"{prefix}r: {run.Properties}");

        foreach (var item in run.Elements)
        {
            DumpRunElement(item, level + 1);
        }
    }

    private static void DumpRunElement(TextElement item, int level)
    {
        string prefix = new string(' ', level * 3);

        string text = "";
        if (item is Text t)
        {
            text = "\"" + t.Value + "\"";
        }
        else if (item is Break b)
        {
            text = b.Type.ToString()+"Break";
        }
        else if (item is CommentRef cr)
        {
            text = $"{cr.Id} - {string.Join(". ", cr.Comment.Paragraphs.Select(p => p.Text))}";
        }
        else if (item is Drawing d)
        {
            if (d.Picture != null)
            {
                var p = d.Picture;
                text = $"{p.FileName} ({Math.Round(p.Width,0)}x{Math.Round(p.Height,0)}) - {p.Name}: \"{p.Description}\"";

                string fn = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    p.FileName);

                /*
                using var input = p.Image.OpenStream();
                using var output = File.OpenWrite(fn);
                input.CopyTo(output);
                */
            }
        }

        Console.WriteLine($"{prefix}{item.Name}: {text}");
    }
}