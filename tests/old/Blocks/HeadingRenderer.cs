using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class HeadingRenderer : DocxObjectRenderer<HeadingBlock>
    {
        public override void Write(DocxRenderer renderer, HeadingBlock obj)
        {
            renderer.WriteChildren(obj.Inline);
            renderer.EndParagraph();
        }
    }
}