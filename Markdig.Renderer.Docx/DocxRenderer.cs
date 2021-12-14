using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Renderer.Docx.Inlines;
using Markdig.Syntax;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using TripleColonInlineRenderer = Markdig.Renderer.Docx.Inlines.TripleColonInlineRenderer;

namespace Markdig.Renderer.Docx
{
    /// <summary>
    /// DoxC renderer for a Markdown <see cref="MarkdownDocument"/> object.
    /// </summary>
    public class DocxObjectRenderer : IDocxRenderer
    {
        private readonly IDocument document;
        readonly List<IDocxObjectRenderer> Renderers;

        public DocxObjectRenderer(IDocument document, string moduleFolder)
        {
            ModuleFolder = moduleFolder;
            this.document = document;

            Renderers = new List<IDocxObjectRenderer>
            {
                // Block handlers
                new HeadingRenderer(),
                new ParagraphRenderer(),
                new ListRenderer(),
                new QuoteBlockRenderer(),
                new QuoteSectionNoteRenderer(),
                new FencedCodeBlockRenderer(),
                new TripleColonRenderer(),
                new FencedCodeBlockRenderer(),
                new TableRenderer(),

                // Inline handlers
                new LiteralInlineRenderer(),
                new EmphasisInlineRenderer(),
                new LineBreakInlineRenderer(),
                new LinkInlineRenderer(),
                new AutolinkInlineRenderer(),
                new CodeInlineRenderer(),
                new DelimiterInlineRenderer(),
                new HtmlEntityInlineRenderer(),
                new LinkReferenceDefinitionRenderer(),
                new TaskListRenderer(),
                new HtmlInlineRenderer(),
                new TripleColonInlineRenderer()
            };
        }

        public string ModuleFolder { get; private set; }
        public Syntax.Block LastBlock { get; private set; }

        public IDocxObjectRenderer FindRenderer(MarkdownObject obj)
        {
            var renderer = Renderers.FirstOrDefault(r => r.CanRender(obj));
#if DEBUG
            if (renderer == null)
            {
                Console.WriteLine($"Missing renderer for {obj.GetType()}");
            }
#endif
            return renderer;
        }

        public void Render(MarkdownDocument markdownDocument)
        {
            for (var index = 0; index < markdownDocument.Count; index++)
            {
                var block = markdownDocument[index];

                // Special case RowBlock and children to generate a full table.
                if (block is RowBlock)
                {
                    var rows = new List<RowBlock>();
                    do
                    {
                        rows.Add((RowBlock) block);
                        block = markdownDocument[++index];
                    } while (block is RowBlock);

                    new RowBlockRenderer().Write(this, document, rows);
                }

                // Find the renderer and process.
                var renderer = FindRenderer(block);
                renderer?.Write(this, document, null, block);
                LastBlock = block;
            }
        }

        public Picture InsertImage(Paragraph currentParagraph, string imageSource, string altText)
        {
            string path = ResolvePath(ModuleFolder, imageSource);
            if (File.Exists(path))
            {
                var img = System.Drawing.Image.FromFile(path);
                var width = img.Width;
                var height = img.Height;

                int finalWidth = width;
                int finalHeight = height;

                if (finalWidth > finalHeight)
                {
                    if (finalWidth > 400)
                    {
                        finalWidth = 400;
                        finalHeight = (int)(400 * ((double)height / width));
                    }
                }
                else
                {
                    if (finalHeight > 400)
                    {
                        finalHeight = 400;
                        finalWidth = (int)(400 * ((double)width / height));
                    }
                }

                var image = document.AddImage(path);
                var picture = image.CreatePicture(imageSource, altText, finalWidth, finalHeight);
                currentParagraph.Append(picture);

                return picture;
            }

            return null;
        }

        private static string ResolvePath(string rootFolder, string path)
            => path.Contains(':')
                ? path
                : Path.IsPathRooted(path)
                    ? path
                    : Path.Combine(rootFolder, path);

    }
}
