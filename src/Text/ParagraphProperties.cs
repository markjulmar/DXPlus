using System;
using System.ComponentModel;
using System.Drawing;
using System.Xml.Linq;
using DXPlus.Helpers;

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
            return value != null ? Math.Round(double.Parse(value) / 20.0, 2) : null;
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
                Xml.GetOrAddElement(Name.Spacing)
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
        /// Set the left indentation in 1/20th pt for this FirstParagraph.
        /// </summary>
        public double LeftIndent
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
        /// Set the right indentation in 1/20th pt for this FirstParagraph.
        /// </summary>
        public double RightIndent
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
        /// The shade pattern applied to this paragraph
        /// </summary>
        public ShadePattern? ShadePattern
        {
            get => Enum.TryParse<ShadePattern>(Xml.Element(Namespace.Main + "shd")?.GetVal(), out var sp) ? sp : null;

            set
            {
                var e = Xml.Element(Namespace.Main + "shd");
                if (value == null)
                {
                    if (e == null) return;
                    value = DXPlus.ShadePattern.Clear;
                }
                e ??= HelperFunctions.CreateDefaultShadeElement(Xml);
                e.SetAttributeValue(Name.MainVal, value.GetEnumName());
            }
        }

        /// <summary>
        /// Shade color used with pattern - use Color.Empty for "auto"
        /// </summary>
        public Color? ShadeColor
        {
            get
            {
                var color = Xml.Element(Namespace.Main + "shd")?.AttributeValue(Name.Color);
                return string.IsNullOrEmpty(color) ? null :
                    color.ToLower() == "auto" ? Color.Empty : ColorTranslator.FromHtml($"#{color}");
            }

            set
            {
                var e = Xml.Element(Namespace.Main + "shd");
                if (value == null)
                {
                    if (e == null) return;
                    value = Color.Empty;
                }

                e ??= HelperFunctions.CreateDefaultShadeElement(Xml);
                e.SetAttributeValue(Name.Color, value == Color.Empty ? "auto" : value.Value.ToHex());
            }
        }

        /// <summary>
        /// Shade fill - use Color.Empty for "auto"
        /// </summary>
        public Color? ShadeFill
        {
            get
            {
                var color = Xml.Element(Namespace.Main + "shd")?.AttributeValue(Namespace.Main + "fill");
                return string.IsNullOrEmpty(color) ? null :
                    color.ToLower() == "auto" ? Color.Empty : ColorTranslator.FromHtml($"#{color}");
            }

            set
            {
                var e = Xml.Element(Namespace.Main + "shd");
                if (value == null)
                {
                    if (e == null) return;
                    value = Color.Empty;
                }

                e ??= HelperFunctions.CreateDefaultShadeElement(Xml);
                e.SetAttributeValue(Namespace.Main + "fill", value == Color.Empty ? "auto" : value.Value.ToHex());
            }
        }

        /// <summary>
        /// Paragraph border
        /// </summary>
        private XElement pBdr => Xml.Element(Namespace.Main + "pBdr");

        /// <summary>
        /// Top border for this paragraph
        /// </summary>
        public BorderEdge TopBorder
        {
            get
            {
                var e = pBdr?.Element(Namespace.Main + BorderEdgeType.Top.GetEnumName());
                return e == null ? null : new BorderEdge(e);
            }
        }

        /// <summary>
        /// Bottom border for this paragraph
        /// </summary>
        public BorderEdge BottomBorder
        {
            get
            {
                var e = pBdr?.Element(Namespace.Main + BorderEdgeType.Bottom.GetEnumName());
                return e == null ? null : new BorderEdge(e);
            }
        }

        /// <summary>
        /// Left border for this paragraph
        /// </summary>
        public BorderEdge LeftBorder
        {
            get
            {
                var e = pBdr?.Element(Namespace.Main + BorderEdgeType.Left.GetEnumName());
                return e == null ? null : new BorderEdge(e);
            }
        }

        /// <summary>
        /// Right border for this paragraph
        /// </summary>
        public BorderEdge RightBorder
        {
            get
            {
                var e = pBdr?.Element(Namespace.Main + BorderEdgeType.Right.GetEnumName());
                return e == null ? null : new BorderEdge(e);
            }
        }

        /// <summary>
        /// Between border for this paragraph
        /// </summary>
        public BorderEdge BetweenBorder
        {
            get
            {
                var e = pBdr?.Element(Namespace.Main + BorderEdgeType.Between.GetEnumName());
                return e == null ? null : new BorderEdge(e);
            }
        }

        /// <summary>
        /// Set all outside edges for the border
        /// </summary>
        public void SetBorders(BorderStyle style, Color color, double? spacing = 1, double size = 2, bool shadow = false)
        {
            if (size is < 2 or > 96)
                throw new ArgumentOutOfRangeException(nameof(size));

            SetBorder(BorderEdgeType.Left, style, color, spacing, size, shadow);
            SetBorder(BorderEdgeType.Top, style, color, spacing, size, shadow);
            SetBorder(BorderEdgeType.Right, style, color, spacing, size, shadow);
            SetBorder(BorderEdgeType.Bottom, style, color, spacing, size, shadow);
        }

        /// <summary>
        /// Set a specific border edge.
        /// </summary>
        /// <exception cref="InvalidEnumArgumentException"></exception>
        public void SetBorder(BorderEdgeType type, BorderStyle style, Color color, double? spacing = 1, double size = 2, bool shadow = false)
        {
            if (size is < 2 or > 96)
                throw new ArgumentOutOfRangeException(nameof(Size));

            if (!Enum.IsDefined(typeof(BorderEdgeType), type))
                throw new InvalidEnumArgumentException(nameof(type), (int)type, typeof(BorderEdgeType));

            var pBdr = Xml.GetOrAddElement(Namespace.Main + "pBdr");

            pBdr.Element(Namespace.Main + type.GetEnumName())?.Remove();

            if (style == BorderStyle.None)
                return;

            var borderXml = new XElement(Namespace.Main + type.GetEnumName(),
                new XAttribute(Name.MainVal, style.GetEnumName()),
                new XAttribute(Name.Size, size));
            if (color != Color.Empty)
                borderXml.Add(new XAttribute(Name.Color, color.ToHex()));
            if (shadow)
                borderXml.Add(new XAttribute(Name.Shadow, true));
            if (spacing != null)
                borderXml.Add(new XAttribute(Namespace.Main + "space", spacing));

            pBdr.Add(borderXml);
        }

        /// <summary>
        /// Get or set the indentation of the first line of this FirstParagraph.
        /// </summary>
        public double FirstLineIndent
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
        /// Get or set the indentation of all but the first line of this FirstParagraph.
        /// </summary>
        public double HangingIndent
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
        /// Change the Right to Left paragraph layout.
        /// </summary>
        public Direction Direction
        {
            get => Xml.Element(Name.RightToLeft) == null ? Direction.LeftToRight : Direction.RightToLeft;
            set => Xml.SetElementValue(Name.RightToLeft, value == Direction.RightToLeft ? string.Empty : null);
        }

        /// <summary>
        /// Default formatting options for the paragraph
        /// </summary>
        public Formatting DefaultFormatting => new(Xml.GetRunProps(true));

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
