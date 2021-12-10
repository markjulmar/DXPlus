using System.Diagnostics;
using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class HeadingRenderer : DocxObjectRenderer<HeadingBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, HeadingBlock heading)
        {
            Debug.Assert(currentParagraph == null);

            currentParagraph = document.AddParagraph();
            switch (heading.Level)
            {
                case 1: currentParagraph.Style(HeadingType.Heading1); break;
                case 2: currentParagraph.Style(HeadingType.Heading2); break;
                case 3: currentParagraph.Style(HeadingType.Heading3); break;
                case 4: currentParagraph.Style(HeadingType.Heading4); break;
                case 5: currentParagraph.Style(HeadingType.Heading5); break;
            }

            WriteChildren(heading, owner, document, currentParagraph);
        }
    }
}