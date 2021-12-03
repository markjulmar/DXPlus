using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class HtmlInlineRenderer : DocxRenderer<HtmlInline>
    {
        protected override void Write(DocxRenderer renderer, HtmlInline obj)
        {
            // HTML inlines are not supported
        }
    }
}