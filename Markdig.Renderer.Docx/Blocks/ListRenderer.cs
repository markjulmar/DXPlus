using System;
using System.Diagnostics;
using System.Linq;
using DXPlus;
using Markdig.Syntax;
using MDTable = Markdig.Extensions.Tables.Table;

namespace Markdig.Renderer.Docx.Blocks
{
    public class ListRenderer : DocxObjectRenderer<ListBlock>
    {
        private int currentLevel = -1;

        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, ListBlock block)
        {
            currentLevel++;
            try
            {
                NumberingDefinition nd;
                if (block.IsOrdered)
                {
                    if (!int.TryParse(block.OrderedStart, out int startNumber))
                        startNumber = 1;
                    nd = document.NumberingStyles.Create(NumberingFormat.Numbered, startNumber);
                }
                else
                {
                    nd = document.NumberingStyles.Create(NumberingFormat.Bullet);
                }

                // ListBlock has a collection of ListItemBlock objects
                // ... which in turn contain paragraphs, tables, code blocks, etc.
                foreach (var listItem in block.Cast<ListItemBlock>())
                {
                    currentParagraph = document.AddParagraph().ListStyle(nd, currentLevel);

                    for (var index = 0; index < listItem.Count; index++)
                    {
                        var childBlock = listItem[index];
                        if (index > 0)
                        {
                            if (childBlock is not MDTable)
                            {
                                // Create a new paragraph to hold this block.
                                // Unless it's a table - that gets appended to the prior paragraph.
                                currentParagraph = currentParagraph.AddParagraph().ListStyle();
                            }
                            else
                            {
                                Console.WriteLine("!");
                            }
                        }

                        Write(childBlock, owner, document, currentParagraph);

                        if (childBlock is MDTable && currentParagraph.Table != null)
                        {
                            currentParagraph.Table.Indent = 36.0;
                        }
                        else if (currentParagraph.Properties.StyleName != "ListParagraph")
                        {
                            currentParagraph.Properties.LeftIndent = 36.0;
                        }
                    }
                }
            }
            finally
            {
                currentLevel--;
            }
        }
    }
}