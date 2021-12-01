using DXPlus;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class QuoteSectionNoteBlockRenderer : DocxObjectRenderer<QuoteSectionNoteBlock>
    {
        protected override void Write(DocxRenderer renderer, QuoteSectionNoteBlock obj)
        {
            if (obj.Count > 0)
            {
                renderer.NewParagraph().Style(HeadingType.IntenseQuote);
                renderer.CurrentParagraph().AppendLine(obj.NoteTypeString);
                renderer.WriteChildren(obj);
                renderer.EndParagraph();
            }
        }
    }
}