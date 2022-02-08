using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using DXPlus;

namespace ReadDocument
{
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

            Console.WriteLine(FormatXml(doc.RawDocument()));
            Console.WriteLine();
            Console.WriteLine(new string('-',10));
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
                string prefix = new(' ', level * 3);
                Console.WriteLine($"{prefix}{ub.Name}");
            }
        }

        private static void DumpParagraph(Paragraph block, int level)
        {
            string prefix = new string(' ', level*3);

            string listInfo = "";
            if (block.IsListItem())
            {
                listInfo = $"{block.GetNumberingFormat()} {block.GetListLevel()} #{block.GetListIndex()+1} ";
            }

            Console.WriteLine($"{prefix}p: {block.Id} StyleName=\"{block.Properties.StyleName}\" {listInfo}{DumpObject(block.Properties.DefaultFormatting)}");
            foreach (var run in block.Runs)
            {
                DumpRun(run, level+1);
            }

            foreach (var comment in block.Comments)
            {
                DumpCommentRef(comment, level+1);
            }
        }

        private static void DumpCommentRef(CommentRange comment, int level)
        {
            string prefix = new string(' ', level * 3);

            Console.WriteLine($"{prefix}Comment id={comment.Comment.Id} by {comment.Comment.AuthorName} ({comment.Comment.AuthorInitials})");
            Console.WriteLine($"{prefix}   > start: {comment.RangeStart.Text}, end: {comment.RangeEnd.Text}");
            foreach (var p in comment.Comment.Paragraphs)
            {
                DumpParagraph(p, level + 1);
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
                    if (d.Hyperlink != null)
                    {
                        text += $", Hyperlink=\"{d.Hyperlink.OriginalString}\"";
                    }
                
                    var p = d.Picture;
                    if (p != null)
                    {
                        text += $"{Environment.NewLine}{prefix}   pic: Id={p.Id}, Rid=\"{p.RelationshipId}\" {p.FileName} ({Math.Round(p.Width,0)}x{Math.Round(p.Height,0)}) - {p.Name}: \"{p.Description}\"";
                        if (p.Hyperlink != null)
                        {
                            text += $", Hyperlink=\"{p.Hyperlink.OriginalString}\"";
                        }

                        if (p.BorderColor != null)
                        {
                            text += $", BorderColor={p.BorderColor}";
                        }

                        string captionText = d.GetCaption();
                        if (captionText != null)
                        {
                            text += $", Caption=\"{captionText}\"";
                        }

                        foreach (var ext in p.ImageExtensions)
                        {
                            if (ext is SvgExtension svg)
                            {
                                text += $"{Environment.NewLine}{prefix}      SvgId={svg.RelationshipId} ({svg.Image.FileName})";
                            }
                            else if (ext is VideoExtension video)
                            {
                                text += $"{Environment.NewLine}{prefix}      Video=\"{video.Source}\" H={video.Height}, W={video.Width}";
                            }
                            else if (ext is DecorativeImageExtension dix)
                            {
                                text += $"{Environment.NewLine}{prefix}      DecorativeImage={dix.Value}";
                            }
                            else if (ext is LocalDpiExtension dpi)
                            {
                                text += $"{Environment.NewLine}{prefix}      LocalDpiOverride={dpi.Value}";
                            }
                            else
                            {
                                text += $"{Environment.NewLine}{prefix}      Extension {ext.UriId}";
                            }
                        }

                            /*
                            string fn;
                            Image theImage;
                            if (p.ImageExtensions.Contains(SvgExtension.ExtensionId))
                            {
                                var svgExt = (SvgExtension) p.ImageExtensions.Get(SvgExtension.ExtensionId);
                                theImage = svgExt.Image;
                                fn = theImage.FileName;
                                text += $", SvgId={svgExt.RelationshipId} ({theImage.FileName})";
                            }
                            else
                            {
                                fn = p.FileName;
                                theImage = p.Image;
                            }

                            fn = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fn);

                            using var input = theImage.OpenStream();
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

            sb.Append(']');
            return sb.ToString();
        }

        public static string FormatXml(string inputXml)
        {
            var document = new XmlDocument();
            document.Load(new StringReader(inputXml));

            var builder = new StringBuilder();
            using (var writer = new XmlTextWriter(new StringWriter(builder)))
            {
                writer.Formatting = System.Xml.Formatting.Indented;
                document.Save(writer);
            }

            return builder.ToString();
        }
    }
}