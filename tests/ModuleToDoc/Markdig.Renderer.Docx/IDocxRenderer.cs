using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx
{
    public interface IDocxRenderer
    {
        Syntax.Block LastBlock { get; }
        IDocxObjectRenderer FindRenderer(MarkdownObject obj);
        void InsertImage(Paragraph currentParagraph, string imageSource, string altText);
    }
}