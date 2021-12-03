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

            if (!literal.Content.IsEmpty)
            {
                currentParagraph.Append(literal.Content.ToString());
            }
        }        
    }
}