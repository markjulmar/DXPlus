using System.Drawing;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class CodeInlineRenderer : DocxRenderer<CodeInline>
    {
        protected override void Write(DocxRenderer renderer, CodeInline obj)
        {
            string code = obj.Content;
            renderer.CurrentParagraph()
                .Append(code)
                .WithFormatting(new Formatting {Font = FontFamily.GenericMonospace, Bold = true});
        }
    }
}