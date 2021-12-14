using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using DXPlus;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class HtmlInlineRenderer : DocxObjectRenderer<HtmlInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, HtmlInline html)
        {
            Debug.Assert(currentParagraph != null);

            string tag = Helpers.GetTag(html.Tag);
            bool isClose = html.Tag.StartsWith("</");
            switch (tag)
            {
                case "kbd":
                    if (isClose)
                    {
                        Run r = currentParagraph.Runs.Last();
                        r.AddFormatting(new Formatting
                        {
                            Bold = true, CapsStyle = CapsStyle.SmallCaps, Font = FontFamily.GenericMonospace,
                            Color = Color.DarkGray
                        });
                    }

                    break;
                case "a":
                    if (!isClose)
                        ProcessRawAnchor(html, owner, document, currentParagraph);
                    break;
                case "br":
                    currentParagraph.AppendLine();
                    break;
                case "rgn":
                    if (!isClose)
                        currentParagraph.Append($"{{rgn {Helpers.ReadLiteralTextAfterTag(html)}}}",
                            new Formatting() { Highlight = Highlight.Cyan });
                    break;
                default:
                    Console.WriteLine($"Encountered unsupported HTML tag: {tag}");
                    break;
            }
        }

        private static void ProcessRawAnchor(HtmlInline html, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
        {
            string text = Helpers.ReadLiteralTextAfterTag(html);
            Regex re = new Regex(@"(?inx)
                <a \s [^>]*
                    href \s* = \s*
                        (?<q> ['""] )
                            (?<url> [^""]+ )
                        \k<q>
                [^>]* >");

            // Ignore if we can't find a URL.
            var m = re.Match(html.Tag);
            if (m.Groups.ContainsKey("url") == false)
            {
                if (text.Length > 0)
                    currentParagraph.Append(text);
            }
            else
            {
                currentParagraph.Append(new Hyperlink(text, new Uri(m.Groups["url"].Value)));
            }
        }
    }
}