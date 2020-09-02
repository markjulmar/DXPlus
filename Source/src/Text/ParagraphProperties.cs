using System;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This manages the [pPr] element in a Word document structure which holds all paragraph-level properties.
    /// </summary>
    public sealed class ParagraphProperties
    {
        private const string DefaultStyle = "Normal";

        internal XElement Xml { get; }

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
                //TODO: should this be case-sensitive?
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
            return value != null ? (double?)Math.Round(double.Parse(value) / 20.0, 2) : null;
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
                   .SetAttributeValue(Namespace.Main + type, Math.Round(value.Value*20.0, 2));
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
        /// Set the left indentation in 1/20th pt for this Paragraph.
        /// </summary>
        public double IndentationLeft
        {
            get => Math.Round(double.Parse(Xml.AttributeValue(Name.Indent, Name.Left) ?? "0") / 20.0, 2);
            set
            {
                if (value == 0)
                {
                    var e = Xml.Element(Name.Indent);
                    e?.Attribute(Name.Left)?.Remove();
                    if (e?.HasAttributes == false)
                        e.Remove();
                }
                else
                {
                    Xml.SetAttributeValue(Name.Indent, Name.Left, Math.Round(value * 20.0, 2));
                }
            }
        }

        /// <summary>
        /// Set the right indentation in 1/20th pt for this Paragraph.
        /// </summary>
        public double IndentationRight
        {
            get => Math.Round(double.Parse(Xml.AttributeValue(Name.Indent, Name.Right) ?? "0") / 20.0, 2);
            set
            {
                if (value == 0)
                {
                    var e = Xml.Element(Name.Indent);
                    e?.Attribute(Name.Right)?.Remove();
                    if (e?.HasAttributes == false)
                        e.Remove();
                }
                else
                {
                    Xml.SetAttributeValue(Name.Indent, Name.Right, Math.Round(value * 20.0, 2));
                }
            }
        }

        /// <summary>
        /// Get or set the indentation of the first line of this Paragraph.
        /// </summary>
        public double IndentationFirstLine
        {
            get => Math.Round(double.Parse(Xml.AttributeValue(Name.Indent, Name.FirstLine) ?? "0") / 20.0, 2);
            set
            {
                var e = Xml.Element(Name.Indent);
                if (value == 0)
                {
                    e?.Attribute(Name.FirstLine)?.Remove();
                    if (e?.HasAttributes == false)
                        e.Remove();
                }
                else
                {
                    e?.Attribute(Name.Hanging)?.Remove();
                    Xml.SetAttributeValue(Name.Indent, Name.FirstLine, Math.Round(value * 20.0, 2));
                }
            }
        }

        /// <summary>
        /// Get or set the indentation of all but the first line of this Paragraph.
        /// </summary>
        public double IndentationHanging
        {
            get => Math.Round(double.Parse(Xml.AttributeValue(Name.Indent, Name.Hanging) ?? "0") / 20.0, 2);

            set
            {
                var e = Xml.Element(Name.Indent);
                if (value == 0)
                {
                    e?.Attribute(Name.Hanging)?.Remove();
                    if (e?.HasAttributes == false)
                        e.Remove();
                }
                else
                {
                    e?.Attribute(Name.FirstLine)?.Remove();
                    Xml.SetAttributeValue(Name.Indent, Name.Hanging, Math.Round(value * 20.0, 2));
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
