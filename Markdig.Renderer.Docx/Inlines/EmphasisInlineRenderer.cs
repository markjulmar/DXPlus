using System.Diagnostics;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class EmphasisInlineRenderer : DocxObjectRenderer<EmphasisInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, EmphasisInline emphasis)
        {
            Debug.Assert(currentParagraph != null);

            // Write children into the paragraph..
            WriteChildren(emphasis, owner, document, currentParagraph);
            // .. and then change the style of that run.
            if (emphasis.DelimiterChar is '*' or '_')
            {
                currentParagraph.WithFormatting(emphasis.DelimiterCount == 2
                    ? new Formatting {Bold = true}
                    : new Formatting {Italic = true});
            }
        }
    }
}