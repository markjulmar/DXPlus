using System.Drawing;
using System.Linq;
using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class CodeBlockRenderer : DocxObjectRenderer<CodeBlock>
    {
        protected override void Write(DocxRenderer renderer, CodeBlock obj)
        {
            var fencedCodeBlock = obj as FencedCodeBlock;
            string language = fencedCodeBlock?.Info ?? string.Empty;
            
            renderer.CurrentParagraph().AppendLine();
            renderer.EndParagraph();
            
            var table = renderer.Document.AddTable(1,1);
            var paragraph = renderer.WriteParagraph(obj);
            paragraph.WithFormatting(new Formatting {Font = FontFamily.GenericMonospace});
            table.Rows[0].Cells[0].AddParagraph(paragraph);
            table.Rows[0].Cells[0].Paragraphs.Where(p => p.Id != paragraph.Id).ToList().ForEach(p => p.Remove());
            table.SetBorders(new TableBorder(TableBorderStyle.Single, BorderSize.One, 10, Color.LightGray));

            renderer.EndParagraph();
            renderer.CurrentParagraph().AppendLine();
        }
    }
}