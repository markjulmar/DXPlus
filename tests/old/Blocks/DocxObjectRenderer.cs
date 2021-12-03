using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public abstract class DocxObjectRenderer<TObject>
        where TObject : MarkdownObject 
    {
        public bool CanRender(MarkdownObject obj) => obj.GetType() == typeof(TObject);
        public abstract void Write(DocxRenderer renderer, TObject obj);
    }
}