using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LineBreakInlineRenderer : DocxObjectRenderer<LineBreakInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LineBreakInline inline)
        {
            //bool hardBreak = inline.IsHard;
            currentParagraph?.AppendLine();
        }
    }
}