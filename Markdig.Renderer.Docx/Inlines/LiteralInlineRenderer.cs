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

            if (literal.Content.IsEmpty) 
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