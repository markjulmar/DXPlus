using DXPlus;
using Markdig.Renderer.Docx.TripleColonExtensions;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class TripleColonRenderer : DocxObjectRenderer<TripleColonBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonBlock block)
        {
            TripleColonProcessor.Write(owner, document, currentParagraph, new TripleColonElement(block));
        }
    }
}
