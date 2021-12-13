using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class QuoteBlockRenderer : DocxObjectRenderer<QuoteBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, QuoteBlock block)
        {
            if (block.Count > 0)
            {
                currentParagraph ??= document.AddParagraph();
                WriteChildren(block, owner, document, currentParagraph.Style(HeadingType.Quote));
            }
        }
    }
}