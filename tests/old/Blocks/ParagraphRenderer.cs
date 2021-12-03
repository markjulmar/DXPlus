using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class ParagraphRenderer : DocxRenderer<ParagraphBlock>
    {
        protected override void Write(BaseRenderer renderer, ParagraphBlock obj)
        {
            renderer.Write(obj.Inline);
        }
    }
}