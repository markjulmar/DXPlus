using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class HtmlEntityInlineRenderer : DocxObjectRenderer<HtmlEntityInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, HtmlEntityInline obj)
        {
            var slice = obj.Transcoded;
            currentParagraph.Append(slice.Text.Substring(slice.Start, slice.Length));
        }
    }
}