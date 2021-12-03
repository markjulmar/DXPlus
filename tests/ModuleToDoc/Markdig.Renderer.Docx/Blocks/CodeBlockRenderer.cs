using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class CodeBlockRenderer : DocxObjectRenderer<CodeBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, CodeBlock block)
        {
            var fencedCodeBlock = block as FencedCodeBlock;
            string language = fencedCodeBlock?.Info;

            if (currentParagraph == null)
                currentParagraph = document.AddParagraph();

            WriteChildren(block, owner, document, currentParagraph);
            currentParagraph.Style("SourceCode");
        }
    }
}