using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx
{
    public interface IDocxRenderer
    {
        string ZonePivot { get; }
        IDocxObjectRenderer FindRenderer(MarkdownObject obj);
        Picture InsertImage(Paragraph currentParagraph, string imageSource, string altText);
    }
}