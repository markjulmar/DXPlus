using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This holds all the formatting properties for a run of text.
    /// This is contained in a (w:rPr) element.
    /// </summary>
    public sealed class Formatting : IEquatable<Formatting>
    {
        internal XElement Xml { get; set; }

        private bool? bold;
        private bool? italic;
        private CapsStyle? capsStyle;
        private Color? color;
        private CultureInfo culture;
        private FontFamily font;
        private double? fontSize;
        private bool? isHidden;
        private Highlight? highlight;
        private int? kerning;
        private Effect? appliedEffect;
        private double? spacing;
        private int? expansionScale;
        private double? position;
        private bool? subscript;
        private bool? superscript;
        private bool? noProof;
        private Strikethrough? strikethrough;
        private Emphasis? emphasis;
        private UnderlineStyle? underlineStyle;
        private Color? underlineColor;

        /// <summary>
        /// Returns whether this paragraph is marked as BOLD
        /// </summary>
        public bool Bold
        {
            get => bold == true;
            set
            {
                bold = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Change the italic state of this paragraph
        /// </summary>
        public bool Italic
        {
            get => italic == true;
            set
            {
                italic = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Change the paragraph to be small caps, capitals or none.
        /// </summary>
        public CapsStyle CapsStyle
        {
            get => capsStyle ?? CapsStyle.None;
            set
            {
                capsStyle = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Returns the applied text color, or None for default.
        /// </summary>
        public Color Color
        {
            get => color ?? Color.Empty;
            set
            {
                color = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Change the culture of the given paragraph.
        /// </summary>
        public CultureInfo Culture
        {
            get => culture;
            set
            {
                culture = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Change the font for the paragraph
        /// </summary>
        public FontFamily Font
        {
            get => font;
            set
            {
                font = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Get or set the font size of this paragraph
        /// </summary>
        public double? FontSize
        {
            get => fontSize;
            set
            {
                fontSize = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// True if this paragraph is hidden.
        /// </summary>
        public bool IsHidden
        {
            get => isHidden == true;
            set
            {
                isHidden = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Gets or sets the highlight on this paragraph
        /// </summary>
        public Highlight Highlight
        {
            get => highlight ?? Highlight.None;
            set
            {
                highlight = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Set the kerning for the paragraph in Pts.
        /// </summary>
        public int? Kerning
        {
            get => kerning;
            set
            {
                kerning = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Applied effect on the paragraph
        /// </summary>
        public Effect Effect
        {
            get => appliedEffect ?? Effect.None;
            set
            {
                appliedEffect = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Specifies the amount of character pitch which shall be added or removed after each character
        /// in this run before the following character is rendered in the document.
        /// </summary>
        public double? Spacing
        {
            get => spacing;
            set
            {
                spacing = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Specifies the amount by which each character shall be expanded or when the character is rendered in the document.
        /// This property stretches or compresses each character in the run.
        /// </summary>
        public int? ExpansionScale
        {
            get => expansionScale;
            set
            {
                if (value < 0 || value > 600)
                    throw new ArgumentOutOfRangeException(nameof(ExpansionScale), "Value must be between 0 and 600.");

                expansionScale = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Specifies the amount (in pts) by which text shall be raised or lowered for this run in relation to the default
        /// baseline of the surrounding non-positioned text. This allows the text to be repositioned without
        /// altering the font size of the contents.
        /// </summary>
        public double? Position
        {
            get => position;
            set
            {
                position = value != null ? (double?) Math.Round(value.Value, 2) : null;
                Save(Xml);
            }
        }

        /// <summary>
        /// Set the paragraph to subscript. Note this is mutually exclusive with Superscript
        /// </summary>
        public bool Subscript
        {
            get => subscript == true;
            set
            {
                subscript = value;
                if (subscript == true && Superscript)
                    superscript = false;
                Save(Xml);
            }
        }

        /// <summary>
        /// Set the paragraph to Superscript. Note this is mutually exclusive with Subscript.
        /// </summary>
        public bool Superscript
        {
            get => superscript == true;
            set
            {
                superscript = value;
                if (superscript == true && Subscript)
                    subscript = false;
                Save(Xml);
            }
        }

        /// <summary>
        /// Specifies that spell/grammar checking is turned off for this text run.
        /// </summary>
        public bool NoProof
        {
            get => noProof == true;
            set
            {
                noProof = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Get or set the underline style for this paragraph
        /// </summary>
        public UnderlineStyle UnderlineStyle
        {
            get => underlineStyle ?? UnderlineStyle.None;
            set
            {
                underlineStyle = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Get or set the underline style for this paragraph
        /// </summary>
        public Color UnderlineColor
        {
            get => underlineColor ?? Color.Empty;
            set
            {
                underlineColor = value;
                underlineStyle ??= UnderlineStyle.SingleLine;
                Save(Xml);
            }
        }

        /// <summary>
        /// Specifies the emphasis mark which shall be displayed for each non-space character in this run.
        /// An emphasis mark is an additional character that is rendered above or below the main character glyph.
        /// </summary>
        public Emphasis Emphasis
        {
            get => emphasis ?? Emphasis.None;
            set
            {
                emphasis = value;
                Save(Xml);
            }
        }

        /// <summary>
        /// Specifies that the text in this paragraph should be displayed with a single or double-line
        /// strikethrough
        /// </summary>
        public Strikethrough StrikeThrough
        {
            get => strikethrough ?? Strikethrough. None;
            set
            {
                strikethrough = value;
                Save(Xml);
            }
        }

        // TODO: add TextBorder (bdr)
        // TODO: add Shading (shd)

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="element">Element containing run properties; if not supplied a new (rPr) element is created.</param>
        public Formatting(XElement element = null)
        {
            Xml = element ?? new XElement(Name.RunProperties);
            Load(Xml);
        }

        /// <summary>
        /// Equals method to compare to another formatting object.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(Formatting other)
        {
            if (ReferenceEquals(this, other))
                return true;

            if (other == null)
                return false;

            return Bold == other.Bold
                   && Italic == other.Italic
                   && CapsStyle == other.CapsStyle
                   && Color == other.Color
                   && Equals(Culture, other.Culture)
                   && Equals(Font, other.Font)
                   && Equals(FontSize, other.FontSize)
                   && IsHidden == other.IsHidden
                   && Highlight == other.Highlight
                   && Kerning == other.Kerning
                   && Effect == other.Effect
                   && Spacing == other.Spacing
                   && ExpansionScale == other.ExpansionScale
                   && Position == other.Position
                   && Subscript == other.Subscript
                   && Superscript == other.Superscript
                   && NoProof == other.NoProof
                   && StrikeThrough == other.StrikeThrough
                   && Emphasis == other.Emphasis
                   && UnderlineStyle == other.UnderlineStyle
                   && UnderlineColor == other.UnderlineColor;
        }

        /// <summary>
        /// Merge in the given formatting into this formatting object. This will add/remove settings
        /// from the given formatting into this one.
        /// </summary>
        /// <param name="other">Formatting to merge in</param>
        public void Merge(Formatting other)
        {
            if (other.bold != null) bold = other.bold;
            if (other.italic != null) italic = other.italic;
            if (other.capsStyle != null) capsStyle = other.capsStyle;
            if (other.color != null) color = other.color;
            if (other.culture != null) culture = other.culture;
            if (other.font != null) font = other.font;
            if (other.fontSize != null) fontSize = other.fontSize;
            if (other.isHidden != null) isHidden = other.isHidden;
            if (other.highlight != null) highlight = other.highlight;
            if (other.kerning != null) kerning = other.kerning;
            if (other.appliedEffect != null) appliedEffect = other.appliedEffect;
            if (other.spacing != null) spacing = other.spacing;
            if (other.expansionScale != null) expansionScale = other.expansionScale;
            if (other.position != null) position = other.position;
            if (other.subscript != null) subscript = other.subscript;
            if (other.superscript != null) superscript = other.superscript;
            if (other.noProof != null) noProof = other.noProof;
            if (other.strikethrough != null) strikethrough = other.strikethrough;
            if (other.emphasis != null) emphasis = other.emphasis;
            if (other.underlineStyle != null) underlineStyle = other.underlineStyle;
            if (other.underlineColor != null) underlineColor = other.underlineColor;
            Save(Xml);
        }

        /// <summary>
        /// Save the formatting to the specified XML block.
        /// </summary>
        /// <param name="xml">rPr element</param>
        private void Save(XElement xml)
        {
            if (xml == null || xml.Name != Name.RunProperties)
                return;

            xml.SetElementValue(Name.Bold, Bold ? string.Empty : null);
            xml.SetElementValue(Name.Bold + "Cs", Bold ? string.Empty : null);

            xml.SetElementValue(Name.Italic, Italic ? string.Empty : null);
            xml.SetElementValue(Name.Italic + "Cs", Italic ? string.Empty : null);

            Xml.Element(Namespace.Main + CapsStyle.SmallCaps.GetEnumName())?.Remove();
            Xml.Element(Namespace.Main + CapsStyle.Caps.GetEnumName())?.Remove();
            if (CapsStyle != CapsStyle.None)
                Xml.Add(new XElement(Namespace.Main + CapsStyle.GetEnumName()));

            Xml.AddElementVal(Name.Color, Color == Color.Empty ? null : Color.ToHex());
            Xml.AddElementVal(Name.Language, Culture?.Name);

            Xml.Element(Name.RunFonts)?.Remove();
            if (Font != null)
            {
                Xml.Add(new XElement(Name.RunFonts,
                    new XAttribute(Namespace.Main + "ascii", Font.Name),
                    new XAttribute(Namespace.Main + "hAnsi", Font.Name),
                    new XAttribute(Namespace.Main + "cs", Font.Name)));
            }

            // Fonts are measured in half-points.
            if (fontSize != null)
            {
                // [0-1638] rounded to nearest half.
                fontSize = Math.Min(Math.Max(0, fontSize.Value), 1638.0);
                fontSize = Math.Round(fontSize.Value * 2, MidpointRounding.AwayFromZero) / 2;

                Xml.AddElementVal(Name.Size, fontSize * 2);
                Xml.AddElementVal(Name.Size + "Cs", fontSize * 2);
            }
            else
            {
                Xml.Element(Name.Size)?.Remove();
                Xml.Element(Name.Size + "Cs")?.Remove();
            }

            Xml.SetElementValue(Name.Vanish, IsHidden ? string.Empty : null);
            Xml.AddElementVal(Name.Highlight, Highlight == Highlight.None ? null : Highlight.GetEnumName());

            // Value is represented in half-pts (1/144") in the document structure.
            Xml.AddElementVal(Name.Kerning, (kerning * 2)?.ToString());

            // Remove all effects first as most are mutually exclusive.
            foreach (var eval in Enum.GetValues(typeof(Effect)).Cast<Effect>())
            {
                if (eval != Effect.None)
                    Xml.Element(Namespace.Main + eval.GetEnumName())?.Remove();
            }

            // Now add the new effect.
            switch (Effect)
            {
                case Effect.None:
                    break;
                case Effect.OutlineShadow:
                    Xml.Add(new XElement(Namespace.Main + Effect.Outline.GetEnumName()),
                        new XElement(Namespace.Main + Effect.Shadow.GetEnumName()));
                    break;
                default:
                    Xml.Add(new XElement(Namespace.Main + Effect.GetEnumName()));
                    break;
            }

            // Represented in 20ths of a pt.
            Xml.AddElementVal(Name.Spacing, spacing == null ? null : Math.Round(spacing.Value * 20.0, 1).ToString(CultureInfo.InvariantCulture));
            Xml.AddElementVal(Namespace.Main + "w", expansionScale?.ToString());
            // Measured in half-pts (1/144")
            Xml.AddElementVal(Name.Position, position != null ? Math.Round(position.Value * 2, 2).ToString(CultureInfo.InvariantCulture) : null);

            if (Subscript)
            {
                Xml.AddElementVal(Name.VerticalAlign, "subscript");
            }
            else if (Superscript)
            {
                Xml.AddElementVal(Name.VerticalAlign, "superscript");
            }
            else
            {
                Xml.Element(Name.VerticalAlign)?.Remove();
            }

            Xml.SetElementValue(Name.NoProof, NoProof ? string.Empty : null);

            Xml.Element(Namespace.Main + Strikethrough.Strike.GetEnumName())?.Remove();
            Xml.Element(Namespace.Main + Strikethrough.DoubleStrike.GetEnumName())?.Remove();
            if (StrikeThrough != Strikethrough.None)
            {
                Xml.AddElementVal(Namespace.Main + StrikeThrough.GetEnumName(), true);
            }

            Xml.AddElementVal(Name.Emphasis, Emphasis != Emphasis.None ? Emphasis.GetEnumName() : null);
            Xml.AddElementVal(Name.Underline, UnderlineStyle != UnderlineStyle.None ? UnderlineStyle.GetEnumName() : null);

            if (UnderlineColor != Color.Empty)
            {
                var e = Xml.GetOrCreateElement(Name.Underline);
                if (e.GetValAttr() == null) // no underline?
                {
                    e.SetAttributeValue(Name.MainVal, UnderlineStyle.SingleLine.GetEnumName());
                }
                e.SetAttributeValue(Name.Color, UnderlineColor.ToHex());
            }
            else
            {
                Xml.Element(Name.Underline)?.SetAttributeValue(Name.Color, null);
            }
        }

        /// <summary>
        /// Load the settings from the supplied XML (rPr) block.
        /// </summary>
        /// <param name="xml">XML to load</param>
        private void Load(XElement xml)
        {
            if (xml == null || xml.Name != Name.RunProperties)
                return;

            bold = xml.Element(Name.Bold) != null ? (bool?)true : null;
            italic = xml.Element(Name.Italic) != null ? (bool?)true : null;

            capsStyle = Xml.Element(Namespace.Main + CapsStyle.SmallCaps.GetEnumName()) != null
                            ? (CapsStyle?)CapsStyle.SmallCaps
                            : Xml.Element(Namespace.Main + CapsStyle.Caps.GetEnumName()) != null
                            ? (CapsStyle?)CapsStyle.Caps
                            : null;

            color = Xml.Element(Name.Color)?.GetValAttr().ToColor() ?? null;

            var name = Xml.Element(Name.Language)?.GetVal();
            culture = name != null ? CultureInfo.GetCultureInfo(name) : null;

            var rfElement = Xml.Element(Name.RunFonts);
            if (rfElement == null)
            {
                font = null;
            }
            else
            {
                name = rfElement.AttributeValue(Namespace.Main + "ascii");
                font = !string.IsNullOrEmpty(name) ? new FontFamily(name) : null;
            }

            fontSize = double.TryParse(Xml.Element(Name.Size)?.GetVal(), out var result) ? (double?)(result / 2) : null;

            isHidden = Xml.Element(Name.Vanish) != null ? (bool?)true : null;
            highlight = Xml.Element(Name.Highlight).TryGetEnumValue<Highlight>(out var hlResult) ? (Highlight?)hlResult : null;
            kerning = int.TryParse(Xml.Element(Name.Kerning)?.GetVal(), out var kernResult) ? (int?)kernResult / 2 : null;

            var appliedEffects = Enum.GetValues(typeof(Effect)).Cast<Effect>()
                .Select(e => Xml.Element(Namespace.Main + e.GetEnumName()))
                .Select(e => (e != null && e.Name.LocalName.TryGetEnumValue<Effect>(out var effResult)) ? effResult : Effect.None)
                .Where(e => e != Effect.None)
                .ToList();

            // The only pair that can be added together.
            if (appliedEffects.Contains(Effect.Outline)
                && appliedEffects.Contains(Effect.Shadow))
            {
                appliedEffects.Remove(Effect.Outline);
                appliedEffects.Remove(Effect.Shadow);
                appliedEffects.Add(Effect.OutlineShadow);
            }

            appliedEffect = appliedEffects.Count == 0 ? null : (DXPlus.Effect?)appliedEffects[0];
            spacing = double.TryParse(Xml.Element(Name.Spacing).GetVal(), out var spacingResult) ? (double?)Math.Round(spacingResult / 20.0, 1) : null;
            expansionScale = int.TryParse(Xml.Element(Namespace.Main + "w").GetVal(), out var scaleResult) ? (int?)scaleResult : null;
            position = double.TryParse(Xml.Element(Name.Position).GetVal(), out var value) ? (double?)value / 2.0 : null;
            subscript = Xml.Element(Name.VerticalAlign)?.GetVal() == "subscript" ? (bool?)true : null;
            superscript = Xml.Element(Name.VerticalAlign)?.GetVal() == "superscript" ? (bool?)true : null;
            noProof = Xml.Element(Name.NoProof) != null ? (bool?)true : null;

            strikethrough = Xml.Element(Namespace.Main + Strikethrough.Strike.GetEnumName()) != null
                ? (Strikethrough?)Strikethrough.Strike
                : Xml.Element(Namespace.Main + Strikethrough.DoubleStrike.GetEnumName()) != null
                    ? (Strikethrough?)Strikethrough.DoubleStrike
                    : null;

            emphasis = Xml.Element(Name.Emphasis).GetVal().TryGetEnumValue<Emphasis>(out var emphResult) ? (Emphasis?)emphResult : null;
            underlineStyle = Xml.Element(Name.Underline).GetVal().TryGetEnumValue<UnderlineStyle>(out var undResult) ? (UnderlineStyle?)undResult : null;

            var attr = Xml.Element(Name.Underline)?.Attribute(Name.Color);
            underlineColor = attr == null || attr.Value == "auto" ? null : (Color?)attr.ToColor();
        }
    }
}
