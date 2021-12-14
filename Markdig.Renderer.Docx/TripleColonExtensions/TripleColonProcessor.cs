using System.ComponentModel;
using System.Drawing;
using DXPlus;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using IContainer = DXPlus.IContainer;

namespace Markdig.Renderer.Docx.TripleColonExtensions
{
    internal static class TripleColonProcessor
    {
        public static void Write(IDocxObjectRenderer renderer, MarkdownObject obj,
            IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonElement extension)
        {
            switch (extension.Extension.Name)
            {
                case "image":
                    HandleImage(owner, document, currentParagraph, extension);
                    break;
                case "zone":
                    HandleZonePivot(renderer, (ContainerBlock)obj, owner, document, currentParagraph, extension);
                    break;
                case "code":
                    HandleCode(renderer, (ContainerBlock) obj, owner, document, currentParagraph, extension);
                    break;
                default:
                    break;
            }
        }

        private static void HandleCode(IDocxObjectRenderer renderer, ContainerBlock containerBlock, IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonElement extension)
        {
            var language = extension.Attributes["language"];
            var source = extension.Attributes["source"];
            extension.Attributes.TryGetValue("range", out var range);
            extension.Attributes.TryGetValue("highlight", out var highlight);

            var p = currentParagraph ?? document.AddParagraph();
            p.Append($"codeBlock: language={language}, source=\"{source}\", range={range}, highlight={highlight}")
                .WithFormatting(new Formatting { Highlight = Highlight.Blue, Color = Color.White });
            if (currentParagraph == null) p.AppendLine();
        }

        private static void HandleZonePivot(IDocxObjectRenderer renderer, ContainerBlock block, 
            IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonElement extension)
        {
            var pivot = extension.Attributes["pivot"];
            if (owner.ZonePivot == null
                || pivot != null && pivot.ToLower().Contains(owner.ZonePivot))
            {
                if (owner.ZonePivot == null)
                {
                    var p = currentParagraph ?? document.AddParagraph();
                    p.Append($"zonePivot: {pivot}")
                        .WithFormatting(new Formatting {Highlight = Highlight.Red, Color = Color.White });
                    if (currentParagraph == null) p.AppendLine();
                }
                
                renderer.WriteChildren(block, owner, document, currentParagraph);
                
                if (owner.ZonePivot == null)
                {
                    var p = currentParagraph ?? document.AddParagraph();
                    p.Append($"end-zonePivot: {pivot}")
                        .WithFormatting(new Formatting { Highlight = Highlight.Red, Color = Color.White });
                    if (currentParagraph == null) p.AppendLine();
                }
            }
        }

        private static void HandleImage(IDocxRenderer owner, IContainer document, 
            Paragraph currentParagraph, TripleColonElement extension)
        {
            currentParagraph ??= document.AddParagraph();

            extension.Attributes.TryGetValue("type", out string type);
            extension.Attributes.TryGetValue("alt-text", out string title);
            extension.Attributes.TryGetValue("source", out string source);
            owner.InsertImage(currentParagraph, source, title);
        }
    }
}
