using System.Drawing;
using DXPlus;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class QuoteSectionNoteRenderer : DocxObjectRenderer<QuoteSectionNoteBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, QuoteSectionNoteBlock block)
        {
            currentParagraph ??= document.AddParagraph();

            if (block.QuoteType == QuoteSectionNoteType.DFMNote 
                && block.NoteTypeString != null)
            {
                currentParagraph
                    .Style(HeadingType.IntenseQuote)
                    .AppendLine(block.NoteTypeString);
                WriteChildren(block, owner, document, currentParagraph);
            }
            else if (block.QuoteType == QuoteSectionNoteType.DFMVideo)
            {
                string videoLink = block.VideoLink;
                currentParagraph.AppendLine($"{{video: {videoLink}}}",
                    new Formatting { Highlight = Highlight.Magenta, Color = Color.White });
            }

            WriteChildren(block, owner, document, currentParagraph);
        }
    }
}
