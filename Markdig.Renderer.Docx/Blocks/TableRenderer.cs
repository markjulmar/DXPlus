using System;
using System.Collections.Generic;
using System.Linq;
using DXPlus;
using Markdig.Extensions.Tables;
using Table = Markdig.Extensions.Tables.Table;

namespace Markdig.Renderer.Docx.Blocks
{
    public class TableRenderer : DocxObjectRenderer<Table>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, Table table)
        {
            bool hasColumnWidth = table.ColumnDefinitions.Any(tableColumnDefinition 
                => tableColumnDefinition.Width != 0.0f && tableColumnDefinition.Width != 1.0f);

            var columnWidths = new List<double>();
            if (hasColumnWidth)
            {
                // Force column widths to be evaluated.
                _ = table.ColumnDefinitions
                    .Select(tableColumnDefinition => Math.Round(tableColumnDefinition.Width * 100) / 100)
                    .ToList();
            }

            // Determine the width of the page
            var section = document.Sections.First();
            double pageWidth = section.Properties.PageWidth - section.Properties.LeftMargin - section.Properties.RightMargin;

            int totalColumns = table.Max(tr => ((TableRow) tr).Count);
            DXPlus.Table documentTable;
            if (currentParagraph != null)
            {
                documentTable = new DXPlus.Table(table.Count, totalColumns);
                currentParagraph.Append(documentTable);
            }
            else
            {
                documentTable = document.AddTable(table.Count, totalColumns);
            }

            bool firstRow = true;

            for (var rowIndex = 0; rowIndex < table.Count; rowIndex++)
            {
                var row = (TableRow) table[rowIndex];
                if (firstRow && row.IsHeader) {
                    documentTable.Design = TableDesign.TableGrid;
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

                    var cellParagraph = documentCell.Paragraphs.First();
                    WriteChildren(cell, owner, document, cellParagraph);
                    
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
                                    cellParagraph.WithProperties(new ParagraphProperties { Alignment = Alignment.Left });
                                    break;
                                case TableColumnAlign.Center:
                                    cellParagraph.WithProperties(new ParagraphProperties { Alignment = Alignment.Center });
                                    break;
                                case TableColumnAlign.Right:
                                    cellParagraph.WithProperties(new ParagraphProperties { Alignment = Alignment.Right });
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
        }
    }
}