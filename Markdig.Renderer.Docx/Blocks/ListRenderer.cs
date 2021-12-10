using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class ListRenderer : DocxObjectRenderer<ListBlock>
    {
        private List currentList;
        private int currentLevel;

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
                    currentList.StartNumber = int.Parse(block.OrderedStart??"1");
            }

            foreach (var item in block.Cast<ListItemBlock>())
            {
                //for (int index = 0; index < item.Count; index++)
                //{
                //    var itemBlock = item[index];

                //    if (itemBlock is LeafBlock)
                //    {
                //        var container = new Paragraph();
                //        Write(itemBlock, owner, document, container);
                //    }
                //    else
                //    {

                //    }
                //}

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