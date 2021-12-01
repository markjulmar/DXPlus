using System;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class AutolinkInlineRenderer : DocxObjectRenderer<AutolinkInline>
    {
        protected override void Write(DocxRenderer renderer, AutolinkInline obj)
        {
            string url = obj.Url;
            renderer.CurrentParagraph().Append(new Hyperlink(url, new Uri(url)));
        }
    }
}