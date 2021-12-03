using System;
using System.IO;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LinkInlineRenderer : DocxObjectRenderer<LinkInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LinkInline link)
        {
            var url = link.GetDynamicUrl?.Invoke() ?? link.Url;

            string title = link.Title;
            if (string.IsNullOrEmpty(title))
            {
                if (link.FirstChild is LiteralInline literal)
                    title = literal.Content.ToString();
            }

            if (link.IsImage)
            {
                owner.InsertImage(currentParagraph, url, title);
            }
            else
            {
                if (string.IsNullOrEmpty(title))
                    title = url;
                currentParagraph.Append(new Hyperlink(title, new Uri(url)));
            }
        }
    }
}