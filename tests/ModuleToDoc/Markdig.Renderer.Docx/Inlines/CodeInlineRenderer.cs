using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class CodeInlineRenderer : DocxObjectRenderer<CodeInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, CodeInline obj)
        {
            string code = obj.Content;
            currentParagraph.Append(code);
            currentParagraph.Style("SourceCode");
                //.WithFormatting(new Formatting {Font = FontFamily.GenericMonospace, Bold = true});
        }
    }
}