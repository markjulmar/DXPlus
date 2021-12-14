using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Markdig.Helpers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx
{
    public static class MarkdigDebug
    {
        public static string Dump(ContainerBlock block, int tabs = 0)
        {
            StringBuilder sb = new StringBuilder();
            DumpContainerBlock(block, tabs, sb);
            return sb.ToString();
        }

        private static void DumpContainerBlock(ContainerBlock cb, int tabs, StringBuilder sb)
        {
            foreach (var item in cb)
            {
                DumpBlock(item, tabs, sb);
            }
        }

        private static void DumpBlock(Block item, int tabs, StringBuilder sb)
        {
            string prefix = new('\t', tabs);

            string details = string.Empty;
            if (item is QuoteSectionNoteBlock qsb)
            {
                if (qsb.VideoLink != null)
                    details = "video: " + qsb.VideoLink;
                else
                    details = qsb.NoteTypeString ?? qsb.SectionAttributeString;
            }
            else if (item is FencedCodeBlock fcb)
                details = $"{fcb.Info} - {fcb.Lines.Count} lines";
            else if (item is TripleColonBlock tcb)
                details = $"{tcb.Extension.Name}: {string.Join($", ", tcb.Attributes.Select(a => $"{a.Key}={a.Value}"))}";

            sb.AppendLine($"{prefix}{item} {details}");

            if (item is LeafBlock pb)
            {
                var inlines = pb.Inline;
                if (inlines != null)
                {
                    foreach (var child in inlines)
                    {
                        DumpInline(child, tabs + 1, sb);
                    }
                }

                if (pb.Lines.Count > 0)
                {
                    foreach (var str in pb.Lines.Cast<StringLine>().Take(pb.Lines.Count))
                        sb.AppendLine($"{prefix} > {str}");
                }
            }
            else if (item is ContainerBlock cb)
            {
                DumpContainerBlock(cb, tabs + 1, sb);
            }
        }

        private static void DumpInline(Inline item, int tabs, StringBuilder sb)
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
                sb.AppendLine($"{prefix}{typeName} {details}");
                DumpContainerInline(cil, tabs + 1, sb);
            }
            else
            {
                sb.AppendLine($"{prefix}{typeName} Value=\"{toString}\" {details}");
            }
        }

        private static void DumpContainerInline(ContainerInline inline, int tabs, StringBuilder sb)
        {
            foreach (var item in inline)
            {
                DumpInline(item, tabs, sb);
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
