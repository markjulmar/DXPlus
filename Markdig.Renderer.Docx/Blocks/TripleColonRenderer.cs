using DXPlus;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class TripleColonRenderer : DocxObjectRenderer<TripleColonBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonBlock block)
        {
            switch (block.Extension.Name)
            {
                case "image":
                    HandleImage(owner, document, currentParagraph, block);
                    break;
                default:
                    break;
            }
        }

        private static void HandleImage(IDocxRenderer owner, IContainer document, Paragraph currentParagraph, TripleColonBlock block)
        {
            currentParagraph ??= document.AddParagraph();

            block.Attributes.TryGetValue("type", out string type);
            block.Attributes.TryGetValue("alt-text", out string title);
            block.Attributes.TryGetValue("source", out string source);
            owner.InsertImage(currentParagraph, source, title);
        }
    }
}
