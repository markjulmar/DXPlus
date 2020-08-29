using System;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    public sealed class ParagraphProperties
    {
        private const string DefaultStyle = "Normal";

        public XElement Xml { get; }

        /// <summary>
        /// True to keep with the next element on the page.
        /// </summary>
        public bool KeepWithNext
        {
            get => Xml.Element(Name.KeepNext) != null;
            set => Xml.SetElementValue(Name.KeepNext, value ? string.Empty : null);
        }

        /// <summary>
        /// Keep all lines in this paragraph together on a page
        /// </summary>
        public bool KeepLinesTogether
        {
            get => Xml.Element(Name.KeepLines) != null;
            set => Xml.SetElementValue(Name.KeepLines, value ? string.Empty : null);
        }


        /// <summary>
        /// Gets or set this Paragraphs text alignment.
        /// </summary>
        public Alignment Alignment
        {
            get => Xml.Element(Name.ParagraphAlignment).GetVal()
                      .TryGetEnumValue<Alignment>(out var result) ? result : Alignment.Left;

            set => Xml.AddElementVal(Name.ParagraphAlignment, value == Alignment.Left ? null : value.GetEnumName());
        }

        ///<summary>
        /// The style name of the paragraph.
        ///</summary>
        public string StyleName
        {
            get
            {
                var styleElement = Xml.Element(Name.ParagraphStyle);
                var attr = styleElement?.Attribute(Name.MainVal);
                return attr != null && !string.IsNullOrEmpty(attr.Value) ? attr.Value : DefaultStyle;
            }
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    value = DefaultStyle;
                Xml.AddElementVal(Name.ParagraphStyle, value != DefaultStyle ? value : null);
            }
        }

        /// <summary>
        /// Set the spacing between lines in this paragraph
        /// </summary>
        public double? LineSpacing
        {
            get => GetLineSpacing("line");
            set => SetLineSpacing("line", value);
        }

        /// <summary>
        /// Set the spacing between lines in this paragraph
        /// </summary>
        public double? LineSpacingAfter
        {
            get => GetLineSpacing("after");
            set => SetLineSpacing("after", value);
        }

        /// <summary>
        /// Set the spacing between lines in this paragraph
        /// </summary>
        public double? LineSpacingBefore
        {
            get => GetLineSpacing("before");
            set => SetLineSpacing("before", value);
        }

        /// <summary>
        /// Helper method to get spacing/xyz
        /// </summary>
        /// <param name="type">type of line spacing to retrieve</param>
        /// <returns>Value or null if not set</returns>
        private double? GetLineSpacing(string type)
        {
            var value = Xml.Element(Name.Spacing).AttributeValue(Namespace.Main + type, null);
            return value != null ? (double?)Math.Round(double.Parse(value) / 20.0, 1) : null;
        }

        /// <summary>
        /// Helper method to set spacing/xyz
        /// </summary>
        /// <param name="type">type of line spacing to adjust</param>
        /// <param name="value">New value</param>
        private void SetLineSpacing(string type, double? value)
        {
            if (value != null)
            {
                Xml.GetOrCreateElement(Name.Spacing)
                   .SetAttributeValue(Namespace.Main + type, Math.Round(value.Value*20.0, 1));
            }
            else
            {
                // Remove the 'spacing' element if it only specifies line spacing
                var el = Xml.Element(Name.Spacing);
                el?.Attribute(Namespace.Main + type)?.Remove();
                if (el?.HasAttributes == false)
                {
                    el.Remove();
                }
            }
        }


        /// <summary>
        /// Formatting properties for the paragraph
        /// </summary>
        /// <param name="xml"></param>
        public ParagraphProperties(XElement xml = null)
        {
            Xml = xml ?? new XElement(Name.ParagraphProperties);
        }
    }
}
