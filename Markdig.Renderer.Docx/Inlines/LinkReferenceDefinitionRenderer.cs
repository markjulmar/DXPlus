using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LinkReferenceDefinitionRenderer : DocxObjectRenderer<LinkReferenceDefinition>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LinkReferenceDefinition linkDef)
        {
            if (!string.IsNullOrEmpty(linkDef.Label))
            {
            }

            if (!string.IsNullOrEmpty(linkDef.Url))
            {
            }

            if (!string.IsNullOrEmpty(linkDef.Title))
            {
            }
        }
    }
}