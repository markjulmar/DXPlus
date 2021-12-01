using System.Diagnostics.CodeAnalysis;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    /// <summary>
    /// An XAML renderer for a <see cref="HeadingBlock"/>.
    /// </summary>
    /// <seealso cref="DocxObjectRenderer{TObject}" />
    public class HeadingRenderer : DocxObjectRenderer<HeadingBlock>
    {
        protected override void Write([NotNull] DocxRenderer renderer, [NotNull] HeadingBlock obj)
        {
            renderer.WriteChildren(obj.Inline);
            renderer.EndParagraph();
        }
    }
}