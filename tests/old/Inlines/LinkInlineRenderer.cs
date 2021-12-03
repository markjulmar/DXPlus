using System;
using System.IO;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LinkInlineRenderer : DocxRenderer<LinkInline>
    {
        protected override void Write(DocxRenderer renderer, LinkInline obj)
        {
            var url = obj.GetDynamicUrl?.Invoke() ?? obj.Url;

            string title = obj.Title;
            if (string.IsNullOrEmpty(title))
            {
                if (obj.FirstChild is LiteralInline literal)
                    title = literal.Content.ToString();
            }
            
            if (obj.IsImage)
            {
                RenderImage(renderer, url, title);
            }
            else
            {
                if (string.IsNullOrEmpty(title))
                    title = url;
                renderer.CurrentParagraph()
                    .Append(new Hyperlink(title, new Uri(url)));
            }
        }

        public static void RenderImage(DocxRenderer renderer, string imageSource, string altText)
        {
            string fullPath = renderer.ResolvePath(imageSource);
            if (File.Exists(fullPath))
            {
                var img = System.Drawing.Image.FromFile(fullPath);
                var width = img.Width;
                var height = img.Height;

                int finalWidth = width;
                int finalHeight = height;

                if (finalWidth > finalHeight)
                {
                    if (finalWidth > 400)
                    {
                        finalWidth = 400;
                        finalHeight = (int)(400 * ((double)height / width));
                    }
                }
                else
                {
                    if (finalHeight > 400)
                    {
                        finalHeight = 400;
                        finalWidth = (int)(400 * ((double)width / height));
                    }
                }
                    
                var image = renderer.Document.AddImage(fullPath);
                var picture = image.CreatePicture(imageSource, altText, finalWidth, finalHeight);
                renderer.CurrentParagraph()
                    .AppendLine()
                    .Append(picture)
                    .AppendLine();
                
                renderer.EndParagraph();
            }
        }
    }
}