using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class ListRenderer : DocxObjectRenderer<ListBlock>
    {
        private List currentList;
        private int currentLevel;
        
        protected override void Write(DocxRenderer renderer, [NotNull] ListBlock listBlock)
        {
            bool isTopList = currentList == null;
            if (!isTopList)
            {
                // First, close off the prior paragraph (if any).
                var existingParagraph = renderer.CurrentParagraph();
                if (existingParagraph != null)
                {
                    currentList.AddItem(existingParagraph, currentLevel);
                    renderer.NewListParagraph();
                }
                
                currentLevel++;
            }
            else
            {
                currentList = new List(listBlock.IsOrdered ? NumberingFormat.Numbered : NumberingFormat.Bulleted);
                if (listBlock.IsOrdered)
                    currentList.StartNumber = int.Parse(listBlock.OrderedStart);
            }
                
            foreach (var item in listBlock.Cast<ListItemBlock>())
            {
                var newParagraph = renderer.WriteParagraph(item);
                if (newParagraph != null)
                {
                    currentList.AddItem(newParagraph, currentLevel);
                }
            }

            if (isTopList)
            {
                renderer.Document.AddList(currentList);
                renderer.EndParagraph();

                currentLevel = 0;
                currentList = null;
            }
        }
    }
}