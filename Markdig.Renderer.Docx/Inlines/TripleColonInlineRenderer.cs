using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Renderer.Docx.TripleColonExtensions;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Inlines
{
    public class TripleColonInlineRenderer : DocxObjectRenderer<TripleColonInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonInline inline)
        {
            TripleColonProcessor.Write(this, inline, owner, document, currentParagraph, new TripleColonElement(inline));
        }
    }
}
