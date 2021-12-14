using System.Diagnostics;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LiteralInlineRenderer : DocxObjectRenderer<LiteralInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LiteralInline literal)
        {
            Debug.Assert(currentParagraph != null);

            // Nothing to render .. ignore.
            if (literal.Content.IsEmpty) 
                return;

            // Surrounded by HTML tags .. ignore.
            if (literal.PreviousSibling is HtmlInline {IsClosed: false} && literal.NextSibling is HtmlInline)
                return;

            if (currentParagraph.Text.Length == 0)
            {
                currentParagraph.SetText(literal.Content.ToString());
            }
            else
            {
                currentParagraph.Append(literal.Content.ToString());
            }
        }        
    }
}