using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

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
    private XElement webVideoPr => Xml.Element(Namespace.WP15 + "webVideoPr")!;
        
    /// <summary>
    /// Height of the embed
    /// </summary>
    public double? Height
    {
        get => double.TryParse(webVideoPr.AttributeValue("h"), out var h) ? h : null;
        set => webVideoPr.SetAttributeValue("h", value);
    }

    /// <summary>
    /// Width of the embed
    /// </summary>
    public double? Width
    {
        get => double.TryParse(webVideoPr.AttributeValue("w"), out var h) ? h : null;
        set => webVideoPr.SetAttributeValue("w", value);
    }

    /// <summary>
    /// The HTML embedded code for the document.
    /// </summary>
    public string? Html
    {
        get => webVideoPr.AttributeValue("embeddedHtml");
        set => webVideoPr.SetAttributeValue("embeddedHtml", value);
    }

    /// <summary>
    /// Attempt to pull the source of the video out of the embed code
    /// </summary>
    public string? Source
    {
        get
        {
            var val = Html?.ToLower();
            if (!string.IsNullOrEmpty(val))
            {
                int pos = val.IndexOf("src=\"", StringComparison.Ordinal);
                if (pos >= 0)
                {
                    pos += 5;
                    int epos = val.IndexOf("\"", pos, StringComparison.Ordinal);
                    if (epos > pos)
                    {
                        return Html?[pos..epos];
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
    public VideoExtension(string url, double width, double height) : base(ExtensionId)
    {
        Xml.Add(new XElement(Namespace.WP15 + "webVideoPr",
            new XAttribute("embeddedHtml", string.Format(VideoText, width, height,url))));
            
        Height = height;
        Width = width;
    }

    /// <summary>
    /// Video embed placeholder
    /// </summary>
    private const string VideoText = "<iframe width=\"{0}\" height=\"{1}\" src=\"{2}\" frameborder=\"0\" allow=\"accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture\" allowfullscreen=\"\" sandbox=\"allow-scripts allow-same-origin allow-popups\"></iframe>";
        
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