using System.Diagnostics;
using System.Drawing;
using System.Linq;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class HtmlInlineRenderer : DocxObjectRenderer<HtmlInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, HtmlInline html)
        {
            Debug.Assert(currentParagraph != null);

            if (html.Tag.StartsWith("</"))
            {
                Run r = currentParagraph.Runs.Last();
                string tag = html.Tag.Substring(2).TrimEnd('>');
                switch (tag.ToLower())
                {
                    case "kbd":
                        r.AddFormatting(new Formatting { Bold = true, CapsStyle = CapsStyle.SmallCaps, Font = FontFamily.GenericMonospace, Color = Color.DarkGray });
                        break;
                }
            }
        }
    }
}