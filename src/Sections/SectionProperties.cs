using System;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This manages a block of section properties (w:sectPr)
    /// </summary>
    public sealed class SectionProperties
    {
        private const int A4Width = 11906;
        private const int A4Height = 16838;

        /// <summary>
        /// XML that makes up this element
        /// </summary>
        internal XElement Xml {get;}

        /// <summary>
        /// Revision id tied to this section
        /// </summary>
        public string RevisionId
        {
            get => Xml.AttributeValue(Namespace.Main + "rsidR");
            set => Xml.SetAttributeValue(Namespace.Main + "rsidR", value);
        }

        /// <summary>
        /// Section break type
        /// </summary>
        public SectionBreakType Type
        {
            get => Xml.Element(Namespace.Main + "type").GetVal()
                .TryGetEnumValue<SectionBreakType>(out var result)
                ? result : SectionBreakType.DefaultNextPage;
            set => Xml.AddElementVal(Namespace.Main + "type", value.GetEnumName());
        }

        /// <summary>
        /// Indicates that Headers.First should be used on the first page.
        /// If this is FALSE, then Headers.First is not used in the doc.
        /// </summary>
        public bool DifferentFirstPage
        {
            get => Xml.Element(Namespace.Main + "titlePg") != null;

            set
            {
                var titlePg = Xml.Element(Namespace.Main + "titlePg");
                if (titlePg == null && value)
                {
                    Xml.Add(new XElement(Namespace.Main + "titlePg", string.Empty));
                }
                else if (titlePg != null && !value)
                {
                    titlePg.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or Sets the Direction of content in this Paragraph.
        /// </summary>
        public Direction Direction
        {
            get => Xml.Element(Name.RightToLeft) == null ? Direction.LeftToRight : Direction.RightToLeft;

            set
            {
                if (value == Direction.RightToLeft)
                {
                    Xml.GetOrCreateElement(Name.RightToLeft);
                }
                else
                {
                    Xml.Element(Name.RightToLeft)?.Remove();
                }
            }
        }

        /// <summary>
        /// Page width value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
        /// </summary>
        public double PageWidth
        {
            get
            {
                var pgSz = Xml.Element(Namespace.Main + "pgSz");
                var w = pgSz?.Attribute(Namespace.Main + "w");
                return w != null && double.TryParse(w.Value, out var value) ? Math.Round(value / 20.0) : 12240.0 / 20.0;
            }

            set => Xml.GetOrCreateElement(Namespace.Main + "pgSz")
                .SetAttributeValue(Namespace.Main + "w", value * 20.0);
        }

        /// <summary>
        /// Page height value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
        /// </summary>
        public double PageHeight
        {
            get
            {
                var pgSz = Xml.Element(Namespace.Main + "pgSz");
                var w = pgSz?.Attribute(Namespace.Main + "h");
                return w != null && double.TryParse(w.Value, out double value) ? Math.Round(value / 20.0) : 15840.0 / 20.0;
            }

            set => Xml.GetOrCreateElement(Namespace.Main + "pgSz")
                .SetAttributeValue(Namespace.Main + "h", value * 20);
        }

        /// <summary>
        /// Orientation
        /// </summary>
        public Orientation Orientation
        {
            get => Xml.AttributeValue(Namespace.Main + "pgSz", Namespace.Main + "orient")
                .TryGetEnumValue<Orientation>(out var result)
                ? result
                : Orientation.Portrait;

            set
            {
                var pgSz = Xml.GetOrCreateElement(Namespace.Main + "pgSz");
                pgSz.SetAttributeValue(Namespace.Main + "orient", value.GetEnumName());
                if (value == Orientation.Landscape)
                {
                    pgSz.SetAttributeValue(Namespace.Main + "w", A4Height);
                    pgSz.SetAttributeValue(Namespace.Main + "h", A4Width);
                }
                else // if (value == Orientation.Portrait)
                {
                    pgSz.SetAttributeValue(Namespace.Main + "w", A4Width);
                    pgSz.SetAttributeValue(Namespace.Main + "h", A4Height);
                }
            }
        }

        /// <summary>
        /// True to mirror the margins
        /// </summary>
        public bool MirrorMargins
        {
            get => Xml.Element(Namespace.Main + "mirrorMargins") != null;
            set
            {
                if (value)
                {
                    Xml.SetElementValue(Namespace.Main + "mirrorMargins", string.Empty);
                }
                else
                {
                    Xml.Element(Namespace.Main + "mirrorMargins")?.Remove();
                }
            }
        }

        /// <summary>
        /// Bottom margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
        /// </summary>
        public double BottomMargin
        {
            get => GetMarginAttribute(Name.Bottom);
            set => SetMarginAttribute(Name.Bottom, value);
        }

        /// <summary>
        /// Left margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
        /// </summary>
        public double LeftMargin
        {
            get => GetMarginAttribute(Name.Left);
            set => SetMarginAttribute(Name.Left, value);
        }

        /// <summary>
        /// Right margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
        /// </summary>
        public double RightMargin
        {
            get => GetMarginAttribute(Name.Right);
            set => SetMarginAttribute(Name.Right, value);
        }

        /// <summary>
        /// Top margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
        /// </summary>
        public double TopMargin
        {
            get => GetMarginAttribute(Name.Top);
            set => SetMarginAttribute(Name.Top, value);
        }

        /// <summary>
        /// Starting page number
        /// </summary>
        public int? StartPageNumber
        {
            get => int.TryParse(Xml.AttributeValue(Namespace.Main + "pgNumType", Namespace.Main + "start"), out var result)
                    ? (int?) result
                    : null;

            set => Xml.SetAttributeValue(Namespace.Main + "pgNumType", Namespace.Main + "start", value?.ToString());
        }

        /// <summary>
        /// Gets or Sets the vertical alignment.
        /// </summary>
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                var val = Xml.Element(Namespace.Main + "vAlign")?.GetValAttr();
                return val != null && val.TryGetEnumValue<VerticalAlignment>(out var result)
                    ? result
                    : VerticalAlignment.Center;
            }

            set => Xml.AddElementVal(Namespace.Main + "vAlign", value.GetEnumName());
        }

        // TODO: add pgBorders, textDirection

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="xml"></param>
        internal SectionProperties(XElement xml)
        {
            Xml = xml ?? throw new ArgumentNullException(nameof(xml));
        }

        /// <summary>
        /// Get a margin
        /// </summary>
        /// <param name="name">Margin to get</param>
        /// <returns>Value in 1/20th pt.</returns>
        private double GetMarginAttribute(XName name)
        {
            var top = Xml.Element(Namespace.Main + "pgMar")?.Attribute(name);
            return top != null && double.TryParse(top.Value, out var value) ? (int)(value / 20.0) : 0;
        }

        /// <summary>
        /// Set a margin
        /// </summary>
        /// <param name="name">Margin to set</param>
        /// <param name="value">Value in 1/20th pt</param>
        private void SetMarginAttribute(XName name, double value)
        {
            Xml.GetOrCreateElement(Namespace.Main + "pgMar")
                .SetAttributeValue(name, value * 20.0);
        }
    }
}
