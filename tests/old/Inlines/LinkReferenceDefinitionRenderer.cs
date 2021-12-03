using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LinkReferenceDefinitionRenderer : DocxRenderer<LinkReferenceDefinition>
    {
        protected override void Write(DocxRenderer renderer, LinkReferenceDefinition linkDef)
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