using System;
using System.Drawing;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

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
                bool addBorder = false;
                if (link.Parent?.ParentBlock?.Parent is QuoteSectionNoteBlock qsnr)
                {
                    if (qsnr.SectionAttributeString != null
                        && qsnr.SectionAttributeString.Contains("mx-imgBorder"))
                        addBorder = true;
                }

                var picture = owner.InsertImage(currentParagraph, url, title);
                if (addBorder && picture != null)
                {
                    picture.BorderColor = Color.DarkGray;
                }
            }
            else
            {
                if (string.IsNullOrEmpty(title))
                    title = url;

                try
                {
                    currentParagraph.Append(new Hyperlink(title, new Uri(url??"", UriKind.RelativeOrAbsolute)));
                }
                catch
                {
                    currentParagraph.Append($"{title} ({url})");
                }
            }
        }
    }
}