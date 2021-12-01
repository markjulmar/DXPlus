using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class ParagraphRenderer : DocxObjectRenderer<ParagraphBlock>
    {
        protected override void Write(DocxRenderer renderer, ParagraphBlock obj)
        {
            renderer.Write(obj.Inline);
        }
    }
}