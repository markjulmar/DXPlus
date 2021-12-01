using System.Linq;
using DXPlus;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Blocks
{
    public class QuoteBlockRenderer : DocxObjectRenderer<QuoteBlock>
    {
        protected override void Write(DocxRenderer renderer, QuoteBlock obj)
        {
            if (obj.Count > 0)
            {
                renderer.NewParagraph().Style(HeadingType.Quote);
                renderer.WriteChildren(obj);
                renderer.EndParagraph();
            }
        }
    }
}