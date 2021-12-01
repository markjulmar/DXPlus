using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LiteralStyle
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }

        public LiteralStyle(LiteralStyle style)
        {
            if (style != null)
            {
                Bold = style.Bold;
                Italic = style.Italic;
            }
        }
    }
    
    public class EmphasisInlineRenderer : DocxObjectRenderer<EmphasisInline>
    {
        internal static LiteralStyle ActiveStyle;
        
        protected override void Write(DocxRenderer renderer, EmphasisInline obj)
        {
            var oldStyle = ActiveStyle;
            
            ActiveStyle = new LiteralStyle(oldStyle);
            if (obj.DelimiterChar == '*' || obj.DelimiterChar == '_')
            {
                if (obj.DelimiterCount == 2)
                {
                    ActiveStyle.Bold = true;
                }
                else
                {
                    ActiveStyle.Italic = true;
                }
            }
            
            renderer.WriteChildren(obj);

            ActiveStyle = oldStyle;
        }
    }
}