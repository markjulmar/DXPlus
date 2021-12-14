using DXPlus;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class InclusionRenderer : DocxObjectRenderer<InclusionBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, InclusionBlock block)
        {
            // TODO: should we pull this file?
            currentParagraph ??= document.AddParagraph();
            currentParagraph.Append($"{{include \"{block.IncludedFilePath}\" {block.Title}}}", 
                new Formatting { Highlight = Highlight.Yellow });
        }
    }
}
