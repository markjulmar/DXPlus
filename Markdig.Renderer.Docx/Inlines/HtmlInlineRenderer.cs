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

            string tag = GetTag(html.Tag);
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
                    {
                        ProcessRawAnchor(html, owner, document, currentParagraph);
                    }
                    break;
                case "br":
                    currentParagraph.AppendLine();
                    break;
                default:
                    Console.WriteLine($"Encountered unsupported HTML tag: {tag}");
                    break;
            }
        }

        private static void ProcessRawAnchor(HtmlInline html, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
        {
            string text = string.Empty;
            if (!html.IsClosed)
            {
                if (html.NextSibling is LiteralInline li)
                {
                    text = li.Content.ToString();
                }
            }

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

        private static string GetTag(string htmlTag)
        {
            if (string.IsNullOrEmpty(htmlTag))
                return null;

            int startPos = 1;
            if (htmlTag.StartsWith("</"))
                startPos = 2;
            int endPos = startPos;
            while (char.IsLetter(htmlTag[endPos]))
                endPos++;
            return htmlTag.Substring(startPos, endPos - startPos).ToLower();
        }
    }
}