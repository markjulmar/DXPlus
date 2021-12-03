using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LineBreakInlineRenderer : DocxRenderer<LineBreakInline>
    {
        protected override void Write(DocxRenderer renderer, LineBreakInline obj)
        {
            bool hardBreak = obj.IsHard;
            if (hardBreak && renderer.CurrentParagraph() != null)
            {
                renderer.NewParagraph();
            }
        }
    }
}