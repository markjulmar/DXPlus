using DXPlus;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class QuoteSectionNoteRenderer : DocxObjectRenderer<QuoteSectionNoteBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, QuoteSectionNoteBlock block)
        {
            currentParagraph ??= document.AddParagraph();

            if (block.NoteTypeString != null)
            {
                currentParagraph
                    .Style(HeadingType.IntenseQuote)
                    .AppendLine(block.NoteTypeString);
                WriteChildren(block, owner, document, currentParagraph);
            }
            else if (block.SectionAttributeString != null)
            {
                // TODO: capture attribute (class)
            }

            WriteChildren(block, owner, document, currentParagraph);
        }
    }
}
