using System;
using System.Collections.Generic;
using System.IO;
using DXPlus;
using Markdig.Renderer.Docx.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class TripleColonBlockRenderer : DocxRenderer<TripleColonBlock>
    {
        protected override void Write(BaseRenderer renderer, TripleColonBlock block)
        {
            switch (block.Extension.Name)
            {
                case "image":
                    WriteImage(renderer, block.Attributes);
                    break;
                default:
                    break;
            }
        }

        private void WriteImage(BaseRenderer renderer, IDictionary<string, string> blockAttributes)
        {
            if (blockAttributes.TryGetValue("source", out string imageSource)
                && !string.IsNullOrWhiteSpace(imageSource))
            {
                blockAttributes.TryGetValue("alt-text", out string altText);
                LinkInlineRenderer.RenderImage(renderer, imageSource, altText);
            }
        }
    }
}