using System;
using System.Text;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx
{
    internal static class Helpers
    {
        public static bool IsSurroundedByHtml(Inline inline)
        {
            if (inline == null) 
                throw new ArgumentNullException(nameof(inline));

            if (inline.PreviousSibling is LiteralInline)
            {
                while (inline.PreviousSibling is LiteralInline)
                    inline = inline.PreviousSibling;
            }

            if (inline.PreviousSibling is HtmlInline {IsClosed: false} startTag)
            {
                while (inline.NextSibling is LiteralInline)
                    inline = inline.NextSibling;

                if (inline.NextSibling is HtmlInline endTag)
                {
                    return GetTag(startTag.Tag) == GetTag(endTag.Tag);
                }
            }

            return false;
        }


        public static string ReadLiteralTextAfterTag(Inline item)
        {
            StringBuilder sb = new();
            if (!item.IsClosed)
            {
                while (item.NextSibling is LiteralInline li)
                {
                    sb.Append(li.Content.ToString());
                    item = item.NextSibling;
                }
            }

            return sb.ToString();
        }

        public static string GetTag(string htmlTag)
        {
            if (string.IsNullOrEmpty(htmlTag))
                return null;

            int startPos = 1;
            if (htmlTag.StartsWith("</"))
                startPos = 2;
            int endPos = startPos;
            while (char.IsLetter(htmlTag[endPos]))
                endPos++;
            return htmlTag.Substring(startPos, endPos - startPos).ToLower();
        }

    }
}
