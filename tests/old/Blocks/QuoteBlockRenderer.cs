using System.Linq;
using DXPlus;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Blocks
{
    public class QuoteBlockRenderer : DocxRenderer<QuoteBlock>
    {
        protected override void Write(BaseRenderer renderer, QuoteBlock obj)
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