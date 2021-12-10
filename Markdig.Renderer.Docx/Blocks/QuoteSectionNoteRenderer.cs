using DXPlus;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class QuoteSectionNoteRenderer : DocxObjectRenderer<QuoteSectionNoteBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, QuoteSectionNoteBlock block)
        {
            currentParagraph ??= document.AddParagraph();

            currentParagraph
                .Style(HeadingType.IntenseQuote)
                .AppendLine(block.NoteTypeString);
            WriteChildren(block, owner, document, currentParagraph);
        }
    }
}
