using System;
using System.Collections.Generic;
using System.Linq;
using DXPlus;
using Markdig.Extensions.Tables;
using Table = Markdig.Extensions.Tables.Table;

namespace Markdig.Renderer.Docx.Blocks
{
    public class TableRenderer : DocxRenderer<Table>
    {
        protected override void Write(BaseRenderer renderer, Table table)
        {
            bool hasColumnWidth = table.ColumnDefinitions.Any(tableColumnDefinition 
                => tableColumnDefinition.Width != 0.0f && tableColumnDefinition.Width != 1.0f);

            var columnWidths = new List<double>();
            if (hasColumnWidth)
            {
                table.ColumnDefinitions
                    .Select(tableColumnDefinition => Math.Round(tableColumnDefinition.Width * 100) / 100)
                    .ToList();
            }
            
            var section = renderer.Document.Sections.First();

            renderer.CurrentParagraph().AppendLine();

            int totalColumns = table.Max(tr => ((TableRow) tr).Count);
            var documentTable = renderer.Document.AddTable(table.Count, totalColumns);

            // Determine the width of the page
            double pageWidth = section.Properties.PageWidth - section.Properties.LeftMargin - section.Properties.RightMargin;
            bool firstRow = true;

            for (var rowIndex = 0; rowIndex < table.Count; rowIndex++)
            {
                var row = (TableRow) table[rowIndex];
                if (firstRow && row.IsHeader) {
                    documentTable.Design = TableDesign.LightGrid;
                }

                firstRow = false;

                for (int colIndex = 0; colIndex < row.Count; colIndex++)
                {
                    var cell = (TableCell) row[colIndex];
                    var documentCell = documentTable.Rows[rowIndex].Cells[colIndex];

                    if (columnWidths.Count > 0)
                    {
                        documentCell.Width = columnWidths[colIndex] * pageWidth;
                        documentCell.SetMargins(0);
                    }

                    var paragraph = renderer.WriteParagraph(cell);
                    documentCell.Paragraphs[0].InsertParagraphBefore(paragraph);
                    
                    if (table.ColumnDefinitions.Count > 0)
                    {
                        var columnIndex = cell.ColumnIndex < 0 || cell.ColumnIndex >= table.ColumnDefinitions.Count
                            ? colIndex
                            : cell.ColumnIndex;
                        columnIndex = columnIndex >= table.ColumnDefinitions.Count
                            ? table.ColumnDefinitions.Count - 1
                            : columnIndex;
                        var alignment = table.ColumnDefinitions[columnIndex].Alignment;
                        if (alignment.HasValue)
                        {
                            switch (alignment)
                            {
                                case TableColumnAlign.Left:
                                    paragraph.WithProperties(new ParagraphProperties { Alignment = Alignment.Left });
                                    break;
                                case TableColumnAlign.Center:
                                    paragraph.WithProperties(new ParagraphProperties { Alignment = Alignment.Center });
                                    break;
                                case TableColumnAlign.Right:
                                    paragraph.WithProperties(new ParagraphProperties { Alignment = Alignment.Right });
                                    break;
                            }
                        }
                    }
                    
                    if (cell.ColumnSpan != 1)
                    {
                    }

                    if (cell.RowSpan != 1)
                    {
                    }
                }
            }

            if (columnWidths.Count == 0)
                documentTable.AutoFit(AutoFit.Contents);

            renderer.EndParagraph();
            renderer.CurrentParagraph().AppendLine();
        }
    }
}