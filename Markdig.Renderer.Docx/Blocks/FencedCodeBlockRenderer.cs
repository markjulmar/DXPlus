using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class FencedCodeBlockRenderer : DocxObjectRenderer<FencedCodeBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, FencedCodeBlock fencedCodeBlock)
        {
            //string language = fencedCodeBlock?.Info;

            currentParagraph ??= document.AddParagraph();

            WriteChildren(fencedCodeBlock, owner, document, currentParagraph);
            currentParagraph.Style("SourceCode");
        }
    }
}