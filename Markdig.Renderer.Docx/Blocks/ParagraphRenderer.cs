using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class ParagraphRenderer : DocxObjectRenderer<ParagraphBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, ParagraphBlock block)
        {
            currentParagraph ??= document.AddParagraph();

            WriteChildren(block, owner, document, currentParagraph);
        }
   }
}