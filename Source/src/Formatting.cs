using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// A text formatting.
    /// </summary>
    public class Formatting : IComparable
    {
        private double? size;
        private int? percentageScale;
        private int? kerning;
        private int? position;
        private double? spacing;

        /// <summary>
        /// A text formatting.
        /// </summary>
        public Formatting()
        {
            CapsStyle = DXPlus.CapsStyle.None;
            StrikeThrough = DXPlus.StrikeThrough.None;
            Script = DXPlus.Script.None;
            Highlight = DXPlus.Highlight.None;
            UnderlineStyle = DXPlus.UnderlineStyle.None;
            Misc = DXPlus.Misc.None;
            Language = CultureInfo.CurrentCulture;
        }

        /// <summary>
        /// Text language
        /// </summary>
        public CultureInfo Language { get; set; }

        /// <summary>
        /// Returns a new identical instance of Formatting.
        /// </summary>
        public Formatting Clone()
        {
            return new Formatting
            {
                Bold = Bold,
                CapsStyle = CapsStyle,
                FontColor = FontColor,
                FontFamily = FontFamily,
                Hidden = Hidden,
                Highlight = Highlight,
                Italic = Italic,
                Language = Language,
                Misc = Misc,
                kerning = kerning,
                percentageScale = percentageScale,
                position = position,
                Script = Script,
                size = size,
                spacing = spacing,
                StrikeThrough = StrikeThrough,
                UnderlineColor = UnderlineColor,
                UnderlineStyle = UnderlineStyle
            };
        }

        public static Formatting Parse(XElement rPr)
        {
            if (rPr == null)
                return null;

            Formatting formatting = new Formatting();

            // Build up the Formatting object.
            foreach (XElement option in rPr.Elements())
            {
                switch (option.Name.LocalName)
                {
                    case "lang":
                        formatting.Language = new CultureInfo(
                            option.GetVal(null) ??
                            option.AttributeValue(DocxNamespace.Main + "eastAsia", null) ??
                            option.AttributeValue(DocxNamespace.Main + "bidi"));
                        break;

                    case "spacing":
                        formatting.Spacing = double.Parse(option.GetVal()) / 20.0;
                        break;

                    case "position":
                    case "kern":
                        formatting.Position = int.Parse(option.GetVal()) / 2;
                        break;

                    case "sz":
                        formatting.Size = int.Parse(
                            option.GetVal()) / 2;
                        break;

                    case "w":
                        formatting.PercentageScale = int.Parse(
                            option.GetVal());
                        break;

                    case "rFonts":
                        formatting.FontFamily =
                            new FontFamily(
                                option.AttributeValue(DocxNamespace.Main + "cs", null) ??
                                option.AttributeValue(DocxNamespace.Main + "ascii", null) ??
                                option.AttributeValue(DocxNamespace.Main + "hAnsi", null) ??
                                option.AttributeValue(DocxNamespace.Main + "eastAsia"));
                        break;

                    case "color":
                        string color = option.GetVal();
                        formatting.FontColor = ColorTranslator.FromHtml(string.Format("#{0}", color));
                        break;

                    case "vanish":
                        formatting.Hidden = true;
                        break;

                    case "b":
                        formatting.Bold = true;
                        break;

                    case "i":
                        formatting.Italic = true;
                        break;

                    case "u":
                        formatting.UnderlineStyle = option.GetEnumValue<UnderlineStyle>();
                        break;
                }
            }

            return formatting;
        }

        internal XElement Xml
        {
            get
            {
                XElement rPr = new XElement(DocxNamespace.Main + "rPr");

                if (Language != null)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "lang", new XAttribute(DocxNamespace.Main + "val", Language.Name)));
                }

                if (spacing.HasValue)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "spacing", new XAttribute(DocxNamespace.Main + "val", spacing.Value * 20)));
                }

                if (position.HasValue)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "position", new XAttribute(DocxNamespace.Main + "val", position.Value * 2)));
                }

                if (kerning.HasValue)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "kern", new XAttribute(DocxNamespace.Main + "val", kerning.Value * 2)));
                }

                if (percentageScale.HasValue)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "w", new XAttribute(DocxNamespace.Main + "val", percentageScale)));
                }

                if (FontFamily != null)
                {
                    rPr.Add(new XElement(
                            DocxNamespace.Main + "rFonts",
                                new XAttribute(DocxNamespace.Main + "ascii", FontFamily.Name),
                                new XAttribute(DocxNamespace.Main + "hAnsi", FontFamily.Name),
                                new XAttribute(DocxNamespace.Main + "cs", FontFamily.Name)
                        )
                    );
                }

                if (Hidden == true)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "vanish"));
                }

                if (Bold == true)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "b"));
                }

                if (Italic == true)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "i"));
                }

                if (UnderlineStyle.HasValue && UnderlineStyle != DXPlus.UnderlineStyle.None)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "u", new XAttribute(DocxNamespace.Main + "val", UnderlineStyle.Value.GetEnumName())));
                }

                if (UnderlineColor.HasValue)
                {
                    // If an underlineColor has been set but no underlineStyle has been set
                    if (UnderlineStyle == DXPlus.UnderlineStyle.None)
                    {
                        UnderlineStyle = DXPlus.UnderlineStyle.SingleLine;
                        rPr.Add(new XElement(DocxNamespace.Main + "u", new XAttribute(DocxNamespace.Main + "val", UnderlineStyle.Value.GetEnumName())));
                    }

                    rPr.Element(DocxNamespace.Main + "u").Add(new XAttribute(DocxNamespace.Main + "color", UnderlineColor.Value.ToHex()));
                }

                if (StrikeThrough.HasValue && StrikeThrough != DXPlus.StrikeThrough.None)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + StrikeThrough.GetEnumName()));
                }

                if (Script.HasValue && Script != DXPlus.Script.None)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "vertAlign", new XAttribute(DocxNamespace.Main + "val", Script.GetEnumName())));
                }

                if (size.HasValue)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "sz", new XAttribute(DocxNamespace.Main + "val", size * 2)));
                    rPr.Add(new XElement(DocxNamespace.Main + "szCs", new XAttribute(DocxNamespace.Main + "val", size * 2)));
                }

                if (FontColor.HasValue)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "color", new XAttribute(DocxNamespace.Main + "val", FontColor.Value.ToHex())));
                }

                if (Highlight.HasValue && Highlight != DXPlus.Highlight.None)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + "highlight", new XAttribute(DocxNamespace.Main + "val", Highlight.GetEnumName())));
                }

                if (CapsStyle.HasValue && CapsStyle.Value != DXPlus.CapsStyle.None)
                {
                    rPr.Add(new XElement(DocxNamespace.Main + CapsStyle.GetEnumName()));
                }

                if (Misc.HasValue && Misc != DXPlus.Misc.None)
                {
                    switch (Misc)
                    {
                        case DXPlus.Misc.OutlineShadow:
                            rPr.Add(new XElement(DocxNamespace.Main + "outline"));
                            rPr.Add(new XElement(DocxNamespace.Main + "shadow"));
                            break;

                        default:
                            rPr.Add(new XElement(DocxNamespace.Main + Misc.GetEnumName()));
                            break;
                    }
                }

                return rPr;
            }
        }

        /// <summary>
        /// This formatting will apply Bold.
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        /// This formatting will apply Italic.
        /// </summary>
        public bool? Italic { get; set; }

        /// <summary>
        /// This formatting will apply StrickThrough.
        /// </summary>
        public StrikeThrough? StrikeThrough { get; set; }

        /// <summary>
        /// The script that this formatting should be, normal, superscript or subscript.
        /// </summary>
        public Script? Script { get; set; }

        private const double MAX_SIZE = 1639;

        /// <summary>
        /// The Size of this text, must be between 0 and 1638.
        /// </summary>
        public double? Size
        {
            get => size;

            set
            {
                if (value != null)
                {
                    // Always round to nearest half
                    value = Math.Round(value.Value * 2, MidpointRounding.AwayFromZero) / 2;
                    if (value < 0 || value > MAX_SIZE)
                    {
                        throw new ArgumentException($"Value must be in the range [0-{MAX_SIZE}]", nameof(Size));
                    }
                }

                size = value;
            }
        }

        /// <summary>
        /// Percentage scale.
        /// </summary>
        public int? PercentageScale
        {
            get => percentageScale;

            set
            {
                if (value != null)
                {
                    int[] valid = new int[] { 200, 150, 100, 90, 80, 66, 50, 33 };
                    if (!valid.Contains(value.Value))
                    {
                        throw new ArgumentOutOfRangeException($"Must be one of the values [{string.Join(",", valid.Select(i => i.ToString()))}]", nameof(PercentageScale));
                    }
                }

                percentageScale = value;
            }
        }

        /// <summary>
        /// The Kerning to apply to this text.
        /// </summary>
        public int? Kerning
        {
            get => kerning;

            set
            {
                if (value != null)
                {
                    int[] valid = new int[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
                    if (!valid.Contains(value.Value))
                    {
                        throw new ArgumentOutOfRangeException($"Must be one of the values [{string.Join(",", valid.Select(i => i.ToString()))}]", nameof(Kerning));
                    }
                }

                kerning = value;
            }
        }

        private const int MAX_POSITION_OR_SPACING = 1585;

        /// <summary>
        /// Text position
        /// </summary>
        public int? Position
        {
            get => position;

            set
            {
                if (value < -1 * MAX_POSITION_OR_SPACING || value > MAX_POSITION_OR_SPACING)
                {
                    throw new ArgumentOutOfRangeException($"Value must be in the range -{MAX_POSITION_OR_SPACING} - {MAX_POSITION_OR_SPACING}", nameof(Position));
                }

                position = value;
            }
        }

        /// <summary>
        /// Text spacing
        /// </summary>
        public double? Spacing
        {
            get => spacing;

            set
            {
                if (value != null)
                {
                    // Round to nearest digit
                    value = Math.Round(value.Value, 1);
                    if (value < -1 * MAX_POSITION_OR_SPACING || value > MAX_POSITION_OR_SPACING)
                    {
                        throw new ArgumentOutOfRangeException($"Value must be in the range -{MAX_POSITION_OR_SPACING} - {MAX_POSITION_OR_SPACING}", nameof(Spacing));
                    }
                }

                spacing = value;
            }
        }

        /// <summary>
        /// The colour of the text.
        /// </summary>
        public Color? FontColor { get; set; }

        /// <summary>
        /// Highlight colour.
        /// </summary>
        public Highlight? Highlight { get; set; }

        /// <summary>
        /// The Underline style that this formatting applies.
        /// </summary>
        public UnderlineStyle? UnderlineStyle { get; set; }

        /// <summary>
        /// The underline colour.
        /// </summary>
        public Color? UnderlineColor { get; set; }

        /// <summary>
        /// Misc settings.
        /// </summary>
        public Misc? Misc { get; set; }

        /// <summary>
        /// Is this text hidden or visible.
        /// </summary>
        public bool? Hidden { get; set; }

        /// <summary>
        /// Capitalization style.
        /// </summary>
        public CapsStyle? CapsStyle { get; set; }

        /// <summary>
        /// The font family of this formatting.
        /// </summary>
        public FontFamily FontFamily { get; set; }

        /// <summary>Compares the current instance with another object of the same type and returns an integer that indicates whether the current instance precedes, follows, or occurs in the same position in the sort order as the other object.</summary>
        /// <param name="obj">An object to compare with this instance.</param>
        /// <returns>A value that indicates the relative order of the objects being compared. The return value has these meanings:  
        ///   Value  
        ///   Meaning  
        ///   Less than zero  
        ///   This instance precedes <paramref name="obj" /> in the sort order.  
        ///   Zero  
        ///   This instance occurs in the same position in the sort order as <paramref name="obj" />.  
        ///   Greater than zero  
        ///   This instance follows <paramref name="obj" /> in the sort order.</returns>
        /// <exception cref="ArgumentException">
        ///   <paramref name="obj" /> is not the same type as this instance.</exception>
        public int CompareTo(object obj)
        {
            Formatting other = (Formatting)obj;
            return other.Hidden != Hidden || other.Bold != Bold || other.Italic != Italic
                || other.StrikeThrough != StrikeThrough || other.Script != Script
                || other.Highlight != Highlight || other.size != size
                || other.FontColor != FontColor || other.UnderlineColor != UnderlineColor
                || other.UnderlineStyle != UnderlineStyle || other.Misc != Misc
                || other.CapsStyle != CapsStyle || other.FontFamily != FontFamily
                || other.percentageScale != percentageScale || other.kerning != kerning
                || other.position != position || other.spacing != spacing
                || !other.Language.Equals(Language)
                ? -1
                : 0;
        }

        public override bool Equals(object obj)
        {
            return ReferenceEquals(this, obj)
                || (!(obj is null)
                    && (obj is Formatting f && CompareTo(f) == 0));
        }

        public override int GetHashCode()
        {
            return Xml.ToString().GetHashCode();
        }

        public static bool operator ==(Formatting left, Formatting right)
        {
            return left is null ? right is null : left.Equals(right);
        }

        public static bool operator !=(Formatting left, Formatting right)
        {
            return !(left == right);
        }

        public static bool operator <(Formatting left, Formatting right)
        {
            return left is null
                ? right is object
                : left.CompareTo(right) < 0;
        }

        public static bool operator <=(Formatting left, Formatting right)
        {
            return left is null || left.CompareTo(right) <= 0;
        }

        public static bool operator >(Formatting left, Formatting right)
        {
            return left is object && left.CompareTo(right) > 0;
        }

        public static bool operator >=(Formatting left, Formatting right)
        {
            return left is null
                ? right is null
                : left.CompareTo(right) >= 0;
        }
    }
}