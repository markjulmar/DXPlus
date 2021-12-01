using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class DelimiterInlineRenderer : DocxObjectRenderer<DelimiterInline>
    {
        protected override void Write(DocxRenderer renderer, DelimiterInline obj)
        {
            renderer.CurrentParagraph().Append(obj.ToLiteral());
            renderer.WriteChildren(obj);
        }
    }
}