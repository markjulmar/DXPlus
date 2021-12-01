using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LiteralInlineRenderer : DocxObjectRenderer<LiteralInline>
    {
        protected override void Write(DocxRenderer renderer, LiteralInline obj)
        {
            if (!obj.Content.IsEmpty)
            {
                var headingParent = obj.Parent?.ParentBlock as HeadingBlock; 
                bool isBold = headingParent == null && EmphasisInlineRenderer.ActiveStyle?.Bold == true;
                bool isItalic = headingParent == null && EmphasisInlineRenderer.ActiveStyle?.Italic == true;

                var paragraph = headingParent != null ? renderer.NewParagraph() : renderer.CurrentParagraph();
                paragraph.Append(obj.Content.ToString());
                if (headingParent != null)
                {
                    switch (headingParent.Level)
                    {
                        case 1:
                            paragraph.Style(HeadingType.Heading1);
                            break;
                        case 2:
                            paragraph.Style(HeadingType.Heading2);
                            break;
                        case 3:
                            paragraph.Style(HeadingType.Heading3);
                            break;
                        case 4:
                            paragraph.Style(HeadingType.Heading4);
                            break;
                        case 5:
                            paragraph.Style(HeadingType.Heading5);
                            break;
                    }
                }
                else
                {
                    paragraph.WithFormatting(new Formatting {Bold = isBold, Italic = isItalic});
                }
            }
        }        
    }
}