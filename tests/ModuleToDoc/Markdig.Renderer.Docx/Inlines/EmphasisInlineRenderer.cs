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
            if (emphasis.DelimiterChar == '*' || emphasis.DelimiterChar == '_')
            {
                if (emphasis.DelimiterCount == 2)
                    currentParagraph.WithFormatting(new Formatting { Bold = true });
                else
                    currentParagraph.WithFormatting(new Formatting { Italic = true });
            }
        }
    }
}
