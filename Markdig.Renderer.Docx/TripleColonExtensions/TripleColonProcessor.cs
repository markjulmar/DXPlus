using DXPlus;

namespace Markdig.Renderer.Docx.TripleColonExtensions
{
    internal static class TripleColonProcessor
    {
        public static void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph,
            TripleColonElement extension)
        {
            switch (extension.Extension.Name)
            {
                case "image":
                    HandleImage(owner, document, currentParagraph, extension);
                    break;
                default:
                    break;
            }
        }

        private static void HandleImage(IDocxRenderer owner, IContainer document, 
            Paragraph currentParagraph, TripleColonElement extension)
        {
            currentParagraph ??= document.AddParagraph();

            extension.Attributes.TryGetValue("type", out string type);
            extension.Attributes.TryGetValue("alt-text", out string title);
            extension.Attributes.TryGetValue("source", out string source);
            owner.InsertImage(currentParagraph, source, title);
        }
    }
}
