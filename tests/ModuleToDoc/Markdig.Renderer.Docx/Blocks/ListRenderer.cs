using System.Diagnostics;
using System.Linq;
using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class ListRenderer : DocxObjectRenderer<ListBlock>
    {
        List currentList = null;
        int currentLevel = 0;

        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, ListBlock block)
        {
            bool topList = false;
            if (currentParagraph != null)
            {
                Debug.Assert(currentList != null);
                currentLevel++;
            }
            else
            {
                topList = true;
                currentList = new List(block.IsOrdered ? NumberingFormat.Numbered : NumberingFormat.Bulleted);
                if (block.IsOrdered)
                    currentList.StartNumber = int.Parse(block.OrderedStart);
            }

            foreach (var item in block.Cast<ListItemBlock>())
            {
                currentParagraph = new Paragraph();
                WriteChildren(item, owner, document, currentParagraph);
                currentList.AddItem(currentParagraph, currentLevel);
            }

            if (topList)
            {
                document.AddList(currentList);
                currentLevel = 0;
                currentList = null;
            }
        }
    }
}