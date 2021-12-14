using System.Collections.Generic;
using System.Linq;
using DXPlus;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class RowBlockRenderer : DocxObjectRenderer<RowBlock>
    {
        public void Write(IDocxRenderer owner, IDocument document, List<RowBlock> rows)
        {
            int totalColumns = rows.Max(r => r.Count);

            var documentTable = document.AddTable(rows.Count, totalColumns);
            documentTable.Design = TableDesign.TableNormal;

            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                for (int colIndex = 0; colIndex < row.Count; colIndex++)
                {
                    var cell = (NestedColumnBlock) row[colIndex];
                    var documentCell = documentTable.Rows[rowIndex].Cells[colIndex];

                    var cellParagraph = documentCell.Paragraphs.First();
                    WriteChildren(cell, owner, document, cellParagraph);
                }
            }
            documentTable.AutoFit(AutoFit.Contents);
        }

        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, RowBlock obj)
        {
            // Not used.
            throw new System.NotImplementedException();
        }
    }
}
