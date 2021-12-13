using DXPlus;
using Markdig.Extensions.TaskLists;
using Markdig.Renderer.Docx.Blocks;

namespace Markdig.Renderer.Docx.Inlines
{
    public class TaskListRenderer : DocxObjectRenderer<TaskList>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TaskList taskListEntry)
        {
            currentParagraph.Append(taskListEntry.Checked
                ? "❎"
                : "⬜ ");

        }
    }
}
