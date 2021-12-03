using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class HtmlEntityInlineRenderer : DocxRenderer<HtmlEntityInline>
    {
        protected override void Write(DocxRenderer renderer, HtmlEntityInline obj)
        {
            bool bold = EmphasisInlineRenderer.ActiveStyle?.Bold == true;
            bool italic = EmphasisInlineRenderer.ActiveStyle?.Italic == true;
            var slice = obj.Transcoded;

            renderer.CurrentParagraph()
                .Append(slice.Text.Substring(slice.Start, slice.Length))
                .WithFormatting(new Formatting {Bold = bold, Italic = italic});
        }
    }
}