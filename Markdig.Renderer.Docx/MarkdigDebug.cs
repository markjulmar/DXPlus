using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Markdig.Helpers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx
{
    public static class MarkdigDebug
    {
        public static void Dump(ContainerBlock block, int tabs = 0)
        {
            foreach (var item in block)
            {
                Dump(item, tabs);
            }
        }

        public static void Dump(Block item, int tabs)
        {
            string prefix = new('\t', tabs);

            string details = string.Empty;
            if (item is QuoteSectionNoteBlock qsb)
                details = qsb.NoteTypeString;
            else if (item is FencedCodeBlock fcb)
                details = $"{fcb.Info} - {fcb.Lines.Count} lines";
            else if (item is TripleColonBlock tcb)
                details = $"{tcb.Extension.Name}: {string.Join($", ", tcb.Attributes.Select(a => $"{a.Key}={a.Value}"))}";

            Console.WriteLine($"{prefix}{item} {details}");

            if (item is LeafBlock pb)
            {
                var inlines = pb.Inline;
                if (inlines != null)
                {
                    foreach (var child in inlines)
                    {
                        Dump(child, tabs + 1);
                    }
                }

                if (pb.Lines.Count > 0)
                {
                    foreach (var str in pb.Lines.Cast<StringLine>().Take(pb.Lines.Count))
                        Console.WriteLine($"{prefix} > {str}");
                }
            }
            else if (item is ContainerBlock cb)
            {
                Dump(cb, tabs + 1);
            }
        }

        public static void Dump(Inline item, int tabs)
        {
            string prefix = new('\t', tabs);

            string typeName = item.GetType().Name;
            string details = GetDebuggerDisplay(item);
            string toString = item.ToString();
            if (details == item.ToString())
                details = "";
            else if (toString == item.GetType().ToString() && item is not ContainerInline)
            {
                toString = details;
                details = "";
            }

            if (item is ContainerInline cil)
            {
                Console.WriteLine($"{prefix}{typeName} {details}");
                Dump(cil, tabs + 1);
            }
            else
            {
                Console.WriteLine($"{prefix}{typeName} Value=\"{toString}\" {details}");
            }
        }

        public static void Dump(ContainerInline inline, int tabs)
        {
            foreach (var item in inline)
            {
                Dump(item, tabs);
            }
        }

        private static string GetDebuggerDisplay(object item)
        {
            var attr = item.GetType().GetCustomAttribute<DebuggerDisplayAttribute>();
            string formatText = attr?.Value;
            if (formatText == null) return null;

            var type = item.GetType();

            foreach (Match match in Regex.Matches(formatText, @"{(.*?)}").Where(m => !m.Groups[1].Value.Contains('{')))
            {
                string pn = match.Groups[1].Value;

                string value;
                var pi = type.GetProperty(pn);
                if (pi != null)
                {
                    value = pi.GetValue(item)?.ToString() ?? "(null)";
                }
                else
                {
                    value = type.GetField(pn)?.GetValue(item)?.ToString() ?? "(null)";
                }

                formatText = formatText.Replace(match.Value, value);
            }

            return formatText;
        }
    }
}
