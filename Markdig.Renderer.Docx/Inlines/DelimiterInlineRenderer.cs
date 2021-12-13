using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class DelimiterInlineRenderer : DocxObjectRenderer<DelimiterInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, DelimiterInline obj)
        {
            currentParagraph.Append(obj.ToLiteral());
        }
    }
}