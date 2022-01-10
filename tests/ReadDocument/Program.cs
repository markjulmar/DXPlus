using System;
using System.IO;
using System.Linq;
using System.Text;
using DXPlus;

public static class Program
{
    public static void Main(string[] args)
    {
        string filename = args.FirstOrDefault();
        if (string.IsNullOrEmpty(filename)
            || !File.Exists(filename))
        {
            Console.WriteLine("Missing or invalid filename.");
            return;
        }

        var doc = Document.Load(filename);

        Console.WriteLine(doc.RawDocument());
        Console.WriteLine();

        foreach (var block in doc.Blocks)
        {
            DumpBlock(block, 0);
        }
    }

    private static void DumpBlock(Block block, int level)
    {
        if (block is Paragraph p)
            DumpParagraph(p, level);
        else if (block is Table t)
            DumpTable(t, level);
        else if (block is UnknownBlock ub)
        {
            string prefix = new string(' ', level * 3);
            Console.WriteLine($"{prefix}{ub.Name}");
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

        Console.WriteLine($"{prefix}p: {block.Id} StyleName=\"{block.Properties.StyleName}\" {listInfo}{DumpObject(block.Properties.DefaultFormatting)}");
        foreach (var run in block.Runs)
        {
            DumpRun(run, level+1);
        }
    }

    private static void DumpTable(Table table, int level)
    {
        string prefix = new string(' ', level*3);
        Console.WriteLine($"{prefix}tbl Design={table.Design} {table.CustomTableDesignName} {table.ConditionalFormatting}");
        foreach (var row in table.Rows)
        {
            DumpRow(row, level + 1);
        }
    }

    private static void DumpRow(TableRow row, int level)
    {
        string prefix = new string(' ', level * 3);
        Console.WriteLine($"{prefix}tr");
        foreach (var cell in row.Cells)
        {
            DumpCell(cell, level + 1);
        }
    }

    private static void DumpCell(TableCell cell, int level)
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

        var parent = run.Parent;
        if (parent is Hyperlink hl)
        {
            Console.WriteLine($"{prefix}hyperlink: {hl.Id} <{hl.Uri}> \"{hl.Text}\"");
            prefix += "   ";
            level++;
        }

        Console.WriteLine($"{prefix}r: {DumpObject(run.Properties)}");
        foreach (var item in run.Elements)
        {
            DumpRunElement(item, level + 1);
        }
    }

    private static void DumpRunElement(ITextElement item, int level)
    {
        string prefix = new string(' ', level * 3);

        string text = "";
        switch (item)
        {
            case Text t:
                text = "\"" + t.Value + "\"";
                break;
            case Break b:
                text = b.Type+"Break";
                break;
            case CommentRef cr:
                text = $"{cr.Id} - {string.Join(". ", cr.Comment.Paragraphs.Select(p => p.Text))}";
                break;
            case Drawing d:
            {
                text = $"{prefix}{item.ElementType}: Id={d.Id} ({Math.Round(d.Width,0)}x{Math.Round(d.Height,0)}) - {d.Name}: \"{d.Description}\"";
                var p = d.Picture;
                if (p != null)
                {
                    text += $"{Environment.NewLine}{prefix}   pic: Id={p.Id}, Rid=\"{p.RelationshipId}\" {p.FileName} ({Math.Round(p.Width,0)}x{Math.Round(p.Height,0)}) - {p.Name}: \"{p.Description}\"";
                    if (p.HasRelatedSvg)
                    {
                        text += $", SvgId={p.SvgRelationshipId} ({p.SvgImage.FileName})";
                    }

                    /*
                string fn = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    p.FileName);

                using var input = p.Image.OpenStream();
                using var output = File.OpenWrite(fn);
                input.CopyTo(output);
                */
                }

                Console.WriteLine(text);
                text = null;
                break;
            }
        }

        if (text != null)
            Console.WriteLine($"{prefix}{item.ElementType}: {text}");
    }

    private static string DumpObject(object obj)
    {
        var sb = new StringBuilder();
        Type t = obj.GetType();

        sb.Append($"{t.Name}: [");
        for (var index = 0; index < t.GetProperties().Length; index++)
        {
            var pi = t.GetProperties()[index];
            object val = pi.GetValue(obj);
            if (val != null)
            {
                if (index > 0) sb.Append(", ");
                sb.Append($"{pi.Name}={val}");
            }
        }

        sb.Append("]");
        return sb.ToString();
    }
}