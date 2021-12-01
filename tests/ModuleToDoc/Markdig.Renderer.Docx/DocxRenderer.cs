using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DXPlus;
using Markdig.Renderers;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Renderer.Docx.Inlines;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx
{
    /// <summary>
    /// DoxC renderer for a Markdown <see cref="MarkdownDocument"/> object.
    /// </summary>
    public class DocxRenderer : RendererBase
    {
        private readonly string moduleFolder;
        public IDocument Document { get; }
        private Paragraph currentParagraph;
        private bool writingEmbeddedParagaph;

        public DocxRenderer(IDocument document, string moduleFolder)
        {
            this.moduleFolder = moduleFolder;
            Document = document;
            
            ObjectRenderers.Add(new TableRenderer());
            ObjectRenderers.Add(new CodeBlockRenderer());
            ObjectRenderers.Add(new ListRenderer());
            ObjectRenderers.Add(new HeadingRenderer());
            //ObjectRenderers.Add(new HtmlBlockRenderer());
            ObjectRenderers.Add(new ParagraphRenderer());
            ObjectRenderers.Add(new QuoteBlockRenderer());
            ObjectRenderers.Add(new QuoteSectionNoteBlockRenderer());
            //ObjectRenderers.Add(new ThematicBreakRenderer());
            ObjectRenderers.Add(new TripleColonBlockRenderer());

            // Default inline renderers
            ObjectRenderers.Add(new AutolinkInlineRenderer());
            ObjectRenderers.Add(new CodeInlineRenderer());
            ObjectRenderers.Add(new DelimiterInlineRenderer());
            ObjectRenderers.Add(new EmphasisInlineRenderer());
            ObjectRenderers.Add(new LineBreakInlineRenderer());
            //ObjectRenderers.Add(new HtmlInlineRenderer());
            ObjectRenderers.Add(new HtmlEntityInlineRenderer());
            ObjectRenderers.Add(new LinkInlineRenderer());
            ObjectRenderers.Add(new LiteralInlineRenderer());
            //ObjectRenderers.Add(new LinkReferenceDefinitionRenderer());
        }

        public string ResolvePath(string path)
        {
            // URL?
            if (path.Contains(':'))
                return path;
            
            // Full root?
            if (Path.IsPathRooted(path))
                return path;

            string fullPath = Path.Combine(moduleFolder, "includes", path);
            if (!File.Exists(fullPath))
            {
                Console.WriteLine($"Missing file: {fullPath}");
            }

            return fullPath;
        }
        
        public Paragraph CurrentParagraph() => currentParagraph ??= Document.AddParagraph();

        public Paragraph NewParagraph()
        {
            if (!writingEmbeddedParagaph)
            {
                currentParagraph = Document.AddParagraph();
            }
            return currentParagraph;
        }

        public override object Render(MarkdownObject markdownObject)
        {
            Write(markdownObject);
            return Document;
        }

        public new void WriteChildren(ContainerBlock containerBlock)
        {
            if (containerBlock != null)
            {
                foreach (var block in containerBlock)
                    Write(block);
            }
        }
        public new void WriteChildren(ContainerInline containerInline)
        {
            if (containerInline != null)
            {
                var inline = containerInline.FirstChild;
                while (inline != null)
                {
                    Write(inline);
                    inline = inline.NextSibling;
                }
            }
        }
        public new void Write(MarkdownObject obj)
        {
            if (obj == null)
                return;

            var renderer = ObjectRenderers.FirstOrDefault(
                    testRenderer => testRenderer.Accept(this, obj));

            if (renderer != null)
            {
                renderer.Write(this, obj);
            }
            else if (obj is ContainerBlock containerBlock)
            {
                WriteChildren(containerBlock);
            }
            else if (obj is ContainerInline containerInline)
            {
                WriteChildren(containerInline);
            }
        }

        public void NewPage()
        {
            currentParagraph = null;
            Document.AddPageBreak();
        }

        readonly Stack<Paragraph> stackedParagraphs = new Stack<Paragraph>();

        public void NewListParagraph()
        {
            currentParagraph = new Paragraph();
        }

        public Paragraph WriteParagraph(ContainerBlock item)
        {
            stackedParagraphs.Push(currentParagraph);
            
            try
            {
                currentParagraph = new Paragraph();
                WriteChildren(item);
                return currentParagraph;
            }
            finally
            {
                currentParagraph = stackedParagraphs.Pop();
            }
        }

        public void EndParagraph()
        {
            if (currentParagraph.Owner != null)
            {
                currentParagraph = null;
            }
        }

        public Paragraph WriteParagraph(LeafBlock item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            
            var paragraph = new Paragraph();
            if (item.Lines.Lines != null)
            {
                var lines = item.Lines;
                var slices = lines.Lines;
                for (int i = 0; i < lines.Count; i++)
                {
                    var slice = slices[i].Slice;
                    string text = slice.Text.Substring(slice.Start, slice.Length);
                    if (i < lines.Count - 1)
                        paragraph.AppendLine(text);
                    else
                        paragraph.Append(text);
                }
            }
            return paragraph;
        }
    }
}
