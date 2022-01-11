using System;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This represents the video extension added to a picture.
    /// </summary>
    public sealed class VideoExtension : DrawingExtension
    {
        /// <summary>
        /// GUID for this extension
        /// </summary>
        public static string ExtensionId => "{C809E66F-F1BF-436E-b5F7-EEA9579F0CBA}";

        /// <summary>
        /// The video tag
        /// </summary>
        private XElement webVideoPr => Xml.Element(Namespace.WP15 + "webVideoPr");
        
        /// <summary>
        /// Height of the embed
        /// </summary>
        public int Height
        {
            get => int.TryParse(webVideoPr.AttributeValue("h"), out var h) ? h : 0;
            set => webVideoPr.SetAttributeValue("h", value.ToString());
        }

        /// <summary>
        /// Width of the embed
        /// </summary>
        public int Width
        {
            get => int.TryParse(webVideoPr.AttributeValue("w"), out var h) ? h : 0;
            set => webVideoPr.SetAttributeValue("w", value.ToString());
        }

        /// <summary>
        /// The HTML embedded code for the document.
        /// </summary>
        public string Html
        {
            get => webVideoPr.AttributeValue("embeddedHtml");
            set => webVideoPr.SetAttributeValue("embeddedHtml", value);
        }

        /// <summary>
        /// Attempt to pull the source of the video out of the embed code
        /// </summary>
        public string Source
        {
            get
            {
                string val = Html?.ToLower();
                if (!string.IsNullOrEmpty(val))
                {
                    int pos = val.IndexOf("src=\"", StringComparison.Ordinal);
                    if (pos >= 0)
                    {
                        pos += 5;
                        int epos = val.IndexOf("\"", pos, StringComparison.Ordinal);
                        if (epos > pos)
                        {
                            return Html.Substring(pos, epos - pos);
                        }
                    }
                }

                return null;
            }
        }
        
        /// <summary>
        /// Public constructor for the video HTML embed
        /// </summary>
        /// <param name="url">Video URL</param>
        /// <param name="width">Width</param>
        /// <param name="height">Height</param>
        public VideoExtension(string url, int width, int height) : base(ExtensionId)
        {
            Xml.Add(new XElement(Namespace.WP15 + "webVideoPr",
                new XAttribute("embeddedHtml", string.Format(videoText, width, height,url))));
            
            Height = height;
            Width = width;
        }

        /// <summary>
        /// Default
        /// </summary>
        private const string videoText = "<iframe width=\"{0}\" height=\"{1}\" src=\"{2}\" frameborder=\"0\" allow=\"accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture\" allowfullscreen=\"\" sandbox=\"allow-scripts allow-same-origin allow-popups\"></iframe>";
        
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="xml">XML element</param>
        internal VideoExtension(XElement xml) : base(xml)
        {
            if (xml.AttributeValue("uri") != UriId)
                throw new ArgumentException("Invalid extension tag for Video.");
        }
    }
}