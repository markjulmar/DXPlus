﻿using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// This holds all the formatting properties for a run of text.
    /// This is typically contained in a (rPr) element.
    /// </summary>
    public sealed class Formatting
    {
        public XElement Xml { get; set; }

        /// <summary>
        /// Returns whether this paragraph is marked as BOLD
        /// </summary>
        public bool Bold
        {
            get => Xml.Element(Name.Bold) != null;
            set
            {
                Xml.SetElementValue(Name.Bold, value ? string.Empty : null);
                Xml.SetElementValue(Name.Bold + "Cs", value ? string.Empty : null);
            }
        }

        /// <summary>
        /// Change the italic state of this paragraph
        /// </summary>
        public bool Italic
        {
            get => Xml.Element(Name.Italic) != null;
            set
            {
                Xml.SetElementValue(Name.Italic, value ? string.Empty : null);
                Xml.SetElementValue(Name.Italic + "Cs", value ? string.Empty : null);
            }
        }

        /// <summary>
        /// Change the paragraph to be small caps, capitals or none.
        /// </summary>
        public CapsStyle CapsStyle
        {
            get => Xml.Element(Namespace.Main + CapsStyle.SmallCaps.GetEnumName()) != null
                ? CapsStyle.SmallCaps
                : Xml.Element(Namespace.Main + CapsStyle.Caps.GetEnumName()) != null
                    ? CapsStyle.Caps
                    : CapsStyle.None;
            set
            {
                Xml.Element(Namespace.Main + CapsStyle.SmallCaps.GetEnumName())?.Remove();
                Xml.Element(Namespace.Main + CapsStyle.Caps.GetEnumName())?.Remove();
                if (value != CapsStyle.None)
                    Xml.Add(new XElement(Namespace.Main + value.GetEnumName()));
            }
        }

        /// <summary>
        /// Returns the applied text color, or None for default.
        /// </summary>
        public Color Color
        {
            get => Xml.Element(Name.Color)?.GetValAttr().ToColor() ?? Color.Empty;
            set => Xml.AddElementVal(Name.Color, value == Color.Empty ? null : value.ToHex());
        }

        /// <summary>
        /// Change the culture of the given paragraph.
        /// </summary>
        public CultureInfo Culture
        {
            get
            {
                var name = Xml.Element(Name.Language)?.GetVal();
                return name != null ? CultureInfo.GetCultureInfo(name) : null;
            }

            set => Xml.AddElementVal(Name.Language, value?.Name);
        }

        /// <summary>
        /// Change the font for the paragraph
        /// </summary>
        public FontFamily Font
        {
            get
            {
                var font = Xml.Element(Name.RunFonts);
                if (font == null) return null;
                string name = font.AttributeValue(Namespace.Main + "ascii");
                return !string.IsNullOrEmpty(name) ? new FontFamily(name) : null;
            }
            set
            {
                Xml.Element(Name.RunFonts)?.Remove();
                if (value != null)
                {
                    Xml.Add(new XElement(Name.RunFonts,
                        new XAttribute(Namespace.Main + "ascii", value.Name),
                        new XAttribute(Namespace.Main + "hAnsi", value.Name),
                        new XAttribute(Namespace.Main + "cs", value.Name)));
                }
            }
        }

        /// <summary>
        /// Get or set the font size of this paragraph
        /// </summary>
        public double? FontSize
        {
            get => double.TryParse(Xml.Element(Name.Size)?.GetVal(), out var result)
                ? (double?)(result / 2)
                : null;

            set
            {
                // Fonts are measured in half-points.
                if (value != null)
                {
                    // [0-1638] rounded to nearest half.
                    double fontSize = value.Value;
                    fontSize = Math.Min(Math.Max(0, fontSize), 1638.0);
                    fontSize = Math.Round(fontSize * 2, MidpointRounding.AwayFromZero) / 2;

                    Xml.AddElementVal(Name.Size, fontSize * 2);
                    Xml.AddElementVal(Name.Size + "Cs", fontSize * 2);
                }
                else
                {
                    Xml.Element(Name.Size)?.Remove();
                    Xml.Element(Name.Size + "Cs")?.Remove();
                }
            }
        }

        /// <summary>
        /// True if this paragraph is hidden.
        /// </summary>
        public bool IsHidden
        {
            get => Xml.Element(Name.Vanish) != null;
            set => Xml.SetElementValue(Name.Vanish, value ? string.Empty : null);
        }

        /// <summary>
        /// Gets or sets the highlight on this paragraph
        /// </summary>
        public Highlight Highlight
        {
            get => Xml.Element(Name.Highlight).TryGetEnumValue<Highlight>(out var result) ? result : Highlight.None;
            set => Xml.AddElementVal(Name.Highlight, value == Highlight.None ? null : value.GetEnumName());
        }

        /// <summary>
        /// Set the kerning for the paragraph in Pts.
        /// </summary>
        public int? Kerning
        {
            // Value is represented in half-pts (1/144") in the document structure.
            get => int.TryParse(Xml.Element(Name.Kerning)?.GetVal(), out var result) ? (int?) result / 2 : null;
            set => Xml.AddElementVal(Name.Kerning, (value*2)?.ToString());
        }

        /// <summary>
        /// Applied effect on the paragraph
        /// </summary>
        public Effect Effect
        {
            get
            {
                var appliedEffects = Enum.GetValues(typeof(Effect)).Cast<Effect>()
                    .Select(e => Xml.Element(Namespace.Main + e.GetEnumName()))
                    .Select(e => (e != null && e.Name.LocalName.TryGetEnumValue<Effect>(out var result)) ? result : Effect.None)
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

                return appliedEffects.Count == 0 ? Effect.None : appliedEffects[0];
            }
            set
            {
                // Remove all effects first as most are mutually exclusive.
                foreach (var eval in Enum.GetValues(typeof(Effect)).Cast<Effect>())
                {
                    if (eval != Effect.None)
                        Xml.Element(Namespace.Main + eval.GetEnumName())?.Remove();
                }

                // Now add the new effect.
                switch (value)
                {
                    case Effect.None:
                        break;
                    case Effect.OutlineShadow:
                        Xml.Add(new XElement(Namespace.Main + Effect.Outline.GetEnumName()),
                                new XElement(Namespace.Main + Effect.Shadow.GetEnumName()));
                        break;
                    default:
                        Xml.Add(new XElement(Namespace.Main + value.GetEnumName()));
                        break;
                }
            }
        }

        /// <summary>
        /// Specifies the amount of character pitch which shall be added or removed after each character
        /// in this run before the following character is rendered in the document.
        /// </summary>
        public double? Spacing
        {
            // Represented in 20ths of a pt.
            get => double.TryParse(Xml.Element(Name.Spacing).GetVal(), out var result) ? (double?)Math.Round(result/20.0,1) : null;
            set => Xml.AddElementVal(Name.Spacing, value == null ? null : Math.Round(value.Value*20.0,1).ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Specifies the amount by which each character shall be expanded or when the character is rendered in the document.
        /// This property stretches or compresses each character in the run.
        /// </summary>
        public int? ExpansionScale
        {
            get => int.TryParse(Xml.Element(Namespace.Main + "w").GetVal(), out var result) ? (int?)result : null;

            set
            {
                if (value != null)
                {
                    if (value < 0 || value > 600)
                        throw new ArgumentOutOfRangeException(nameof(ExpansionScale), "Value must be between 0 and 600.");
                }
                Xml.AddElementVal(Namespace.Main + "w", value?.ToString());
            }
        }

        /// <summary>
        /// Specifies the amount (in pts) by which text shall be raised or lowered for this run in relation to the default
        /// baseline of the surrounding non-positioned text. This allows the text to be repositioned without
        /// altering the font size of the contents.
        /// </summary>
        public double? Position
        {
            // Measured in half-pts (1/144")
            get => double.TryParse(Xml.Element(Name.Position).GetVal(), out var value) ? (double?) value/2.0 : null;
            set => Xml.AddElementVal(Name.Position, value != null 
                    ? Math.Round(value.Value * 2, 2).ToString(CultureInfo.InvariantCulture) 
                    : null);
        }

        /// <summary>
        /// Set the paragraph to subscript. Note this is mutually exclusive with Superscript
        /// </summary>
        public bool Subscript
        {
            get => Xml.Element(Name.VerticalAlign)?.GetVal() == "subscript";
            set
            {
                if (value)
                {
                    Xml.AddElementVal(Name.VerticalAlign, "subscript");
                }
                else if (Subscript)
                {
                    Xml.Element(Name.VerticalAlign)?.Remove();
                }
            }
        }

        /// <summary>
        /// Set the paragraph to Superscript. Note this is mutually exclusive with Subscript.
        /// </summary>
        public bool Superscript
        {
            get => Xml.Element(Name.VerticalAlign)?.GetVal() == "superscript";
            set
            {
                if (value)
                {
                    Xml.AddElementVal(Name.VerticalAlign, "superscript");
                }
                else if (Superscript)
                {
                    Xml.Element(Name.VerticalAlign)?.Remove();
                }
            }
        }

        /// <summary>
        /// Specifies that spell/grammar checking is turned off for this text run.
        /// </summary>
        public bool NoProof
        {
            get => Xml.Element(Name.NoProof) != null;
            set => Xml.SetElementValue(Name.NoProof, value ? string.Empty : null);
        }

        /// <summary>
        /// Get or set the underline style for this paragraph
        /// </summary>
        public UnderlineStyle UnderlineStyle
        {
            get => Xml.Element(Name.Underline).GetVal().TryGetEnumValue<UnderlineStyle>(out var result) 
                ? result : UnderlineStyle.None;
            set => Xml.AddElementVal(Name.Underline, value != UnderlineStyle.None ? value.GetEnumName() : null);
        }

        /// <summary>
        /// Get or set the underline style for this paragraph
        /// </summary>
        public Color UnderlineColor
        {
            get
            {
                var attr = Xml.Element(Name.Underline)?.Attribute(Name.Color);
                return attr == null || attr.Value == "auto" ? Color.Empty : attr.ToColor();
            }

            set
            {
                if (value != Color.Empty)
                {
                    var e = Xml.GetOrCreateElement(Name.Underline);
                    if (e.GetValAttr() == null) // no underline?
                    {
                        e.SetAttributeValue(Name.MainVal, UnderlineStyle.SingleLine.GetEnumName());
                    }
                    e.SetAttributeValue(Name.Color, value.ToHex());
                }
                else
                {
                    Xml.Element(Name.Underline)?.SetAttributeValue(Name.Color, null);
                }
            }
        }

        /// <summary>
        /// Specifies the emphasis mark which shall be displayed for each non-space character in this run.
        /// An emphasis mark is an additional character that is rendered above or below the main character glyph.
        /// </summary>
        public Emphasis Emphasis
        {
            get => Xml.Element(Name.Emphasis).GetVal().TryGetEnumValue<Emphasis>(out var result)
                ? result : Emphasis.None;
            set => Xml.AddElementVal(Name.Emphasis, value != Emphasis.None ? value.GetEnumName() : null);
        }

        /// <summary>
        /// Specifies that the text in this paragraph should be displayed with a single or double-line
        /// strikethrough
        /// </summary>
        public Strikethrough StrikeThrough
        {
            get => Xml.Element(Namespace.Main + Strikethrough.Strike.GetEnumName()) != null
                ? Strikethrough.Strike
                : Xml.Element(Namespace.Main + Strikethrough.DoubleStrike.GetEnumName()) != null
                    ? Strikethrough.DoubleStrike
                    : Strikethrough.None;
            set
            {
                Xml.Element(Namespace.Main + Strikethrough.Strike.GetEnumName())?.Remove();
                Xml.Element(Namespace.Main + Strikethrough.DoubleStrike.GetEnumName())?.Remove();

                if (value != Strikethrough.None)
                {
                    Xml.AddElementVal(Namespace.Main + value.GetEnumName(), true);
                }
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
        }
    }
}
