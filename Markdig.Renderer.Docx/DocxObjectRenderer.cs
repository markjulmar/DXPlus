using System.Linq;
using DXPlus;
using Markdig.Helpers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx
{
    public interface IDocxObjectRenderer
    {
        bool CanRender(MarkdownObject obj);
        void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, MarkdownObject obj);
        void WriteChildren(LeafBlock leafBlock, IDocxRenderer owner, IDocument document, Paragraph currentParagraph);
        void WriteChildren(ContainerBlock container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph);
        void WriteChildren(ContainerInline container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph);
        void Write(MarkdownObject item, IDocxRenderer owner, IDocument document, Paragraph currentParagraph);
    }

    public abstract class DocxObjectRenderer<TObject> : IDocxObjectRenderer
        where TObject : MarkdownObject 
    {
        public virtual bool CanRender(MarkdownObject obj) => obj.GetType() == typeof(TObject) || obj is TObject;
        public abstract void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TObject obj);
        void IDocxObjectRenderer.Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, MarkdownObject obj)
            => Write(owner, document, currentParagraph, (TObject)obj);

        public void WriteChildren(LeafBlock leafBlock, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
        {
            var inlines = leafBlock.Inline;
            if (inlines != null)
            {
                foreach (var child in inlines)
                {
                    Write(child, owner, document, currentParagraph);
                }
            }

            if (leafBlock.Lines.Count > 0)
            {
                int index = 0;
                int count = leafBlock.Lines.Count;
                foreach (var text in leafBlock.Lines.Cast<StringLine>().Take(count))
                {
                    if (++index < count)
                        currentParagraph.AppendLine(text.ToString());
                    else
                        currentParagraph.Append(text.ToString());
                }
            }
        }

        public void WriteChildren(ContainerBlock container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
        {
            foreach (var block in container)
            {
                Write(block, owner, document, currentParagraph);
            }
        }

        public void WriteChildren(ContainerInline container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
        {
            foreach (var inline in container)
            {
                Write(inline, owner, document, currentParagraph);
            }
        }

        public void Write(MarkdownObject item, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
        {
            var renderer = owner.FindRenderer(item);
            if (renderer != null)
            {
                renderer.Write(owner, document, currentParagraph, item);
            }
        }
    }
}