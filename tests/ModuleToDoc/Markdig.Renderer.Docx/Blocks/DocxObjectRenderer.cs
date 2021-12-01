using Markdig.Renderers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Blocks
{
    /// <summary>
    /// A base class for XAML rendering <see cref="Block"/> and <see cref="Inline"/> Markdown objects.
    /// </summary>
    /// <typeparam name="TObject">The type of the object.</typeparam>
    /// <seealso cref="IMarkdownObjectRenderer" />
    public abstract class DocxObjectRenderer<TObject> : MarkdownObjectRenderer<DocxRenderer,TObject>
        where TObject : MarkdownObject
    {
    }
}