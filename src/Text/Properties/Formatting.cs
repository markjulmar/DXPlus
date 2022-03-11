using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus;

/// <summary>
/// This holds all the formatting properties for a run of text.
/// This is contained in a (w:rPr) element.
/// </summary>
public sealed class Formatting : IEquatable<Formatting>
{
    internal XElement Xml { get; }
    private readonly HashSet<string> setProperties = new();

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
            setProperties.Add(nameof(Bold));
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
            setProperties.Add(nameof(Italic));
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
            setProperties.Add(nameof(CapsStyle));
        }
    }

    /// <summary>
    /// Returns the applied text color, or None for default.
    /// </summary>
    public ColorValue? Color
    {
        get
        {
            var color = Xml.Element(Name.Color);
            return color == null ? null : new ColorValue(color);
        }

        set
        {
            if (value == null || value.IsEmpty)
            {
                Xml.Element(Name.Color)?.Remove();
            }
            else
            {
                var element = Xml.GetOrAddElement(Name.Color);
                value.SetElementValues(element);
            }
        }
    }

    /// <summary>
    /// Change the culture of the given paragraph.
    /// </summary>
    public CultureInfo? Culture
    {
        get
        {
            var name = Xml.Element(Name.Language)?.GetVal();
            return name != null ? CultureInfo.GetCultureInfo(name) : null;
        }

        set
        {
            Xml.AddElementVal(Name.Language, value?.Name);
            setProperties.Add(nameof(Culture));
        }
    }

    /// <summary>
    /// Change the font for the paragraph
    /// </summary>
    public FontFamily? Font
    {
        get
        {
            var font = Xml.Element(Name.RunFonts);
            if (font == null) return null;
            var name = font.AttributeValue(Namespace.Main + "ascii");
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
            setProperties.Add(nameof(Font));
        }
    }

    /// <summary>
    /// Get or set the font size of this paragraph in points
    /// </summary>
    public double? FontSize
    {
        get => double.TryParse(Xml.Element(Name.Size)?.GetVal(), out var result) ? result / 2 : null;

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
            setProperties.Add(nameof(FontSize));
        }
    }

    /// <summary>
    /// True if this paragraph is hidden.
    /// </summary>
    public bool IsHidden
    {
        get => Xml.Element(Name.Vanish) != null;
        set
        {
            Xml.SetElementValue(Name.Vanish, value ? string.Empty : null);
            setProperties.Add(nameof(IsHidden));
        }
    }

    /// <summary>
    /// Gets or sets the highlight on this paragraph
    /// </summary>
    public Highlight Highlight
    {
        get => Xml.Element(Name.Highlight)?.TryGetEnumValue<Highlight>(out var result) == true ? result : Highlight.None;
        set
        {
            Xml.AddElementVal(Name.Highlight, value == Highlight.None ? null : value.GetEnumName());
            setProperties.Add(nameof(Highlight));
        }
    }

    /// <summary>
    /// Set the kerning for the paragraph in Pts.
    /// </summary>
    public int? Kerning
    {
        // Value is represented in half-pts (1/144") in the document structure.
        get => int.TryParse(Xml.Element(Name.Kerning)?.GetVal(), out var result) ? (int?) result / 2 : null;
        set
        {
            Xml.AddElementVal(Name.Kerning, (value * 2)?.ToString());
            setProperties.Add(nameof(Kerning));
        }
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
            setProperties.Add(nameof(Effect));
        }
    }

    /// <summary>
    /// Specifies the amount of character pitch which shall be added or removed after each character
    /// in this run before the following character is rendered in the document. This value is represented in dxa units.
    /// </summary>
    public double? Spacing
    {
        get => double.TryParse(Xml.Element(Name.Spacing).GetVal(), out var result) ? result : null;
        set
        {
            if (value == null)
            {
                Xml.Attribute(Name.Spacing)?.Remove();
            }
            else
            {
                Xml.AddElementVal(Name.Spacing, value.Value.ToString(CultureInfo.InvariantCulture));
            }

            setProperties.Add(nameof(Spacing));
        }
    }

    /// <summary>
    /// Specifies the amount by which each character shall be expanded or when the character is rendered in the document.
    /// This property stretches or compresses each character in the run.
    /// </summary>
    public int? ExpansionScale
    {
        get => int.TryParse(Xml.Element(Namespace.Main + "w").GetVal(), out var result) ? result : null;

        set
        {
            if (value != null)
            {
                if (value < 0 || value > 600)
                    throw new ArgumentOutOfRangeException(nameof(ExpansionScale), "Value must be between 0 and 600.");
            }
            Xml.AddElementVal(Namespace.Main + "w", value?.ToString());
            setProperties.Add(nameof(ExpansionScale));
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
        set
        {
            Xml.AddElementVal(Name.Position, value != null
                ? Math.Round(value.Value * 2, 2).ToString(CultureInfo.InvariantCulture)
                : null);
            setProperties.Add(nameof(Position));
        }
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
            setProperties.Add(nameof(Subscript));
            setProperties.Add(nameof(Superscript));
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
            setProperties.Add(nameof(Subscript));
            setProperties.Add(nameof(Superscript));
        }
    }

    /// <summary>
    /// Specifies that spell/grammar checking is turned off for this text run.
    /// </summary>
    public bool NoProof
    {
        get => Xml.Element(Name.NoProof) != null;
        set
        {
            Xml.SetElementValue(Name.NoProof, value ? string.Empty : null);
            setProperties.Add(nameof(NoProof));
        }
    }

    /// <summary>
    /// Get or set the underline style for this paragraph
    /// </summary>
    public Underline? Underline
    {
        get
        {
            var e = Xml.Element(Name.Underline);
            return e == null ? null : new(Xml, e);
        }
        set
        {
            Xml.Element(Name.Underline)?.Remove();
            if (value == null) return;
            Xml.Add(value.Xml);
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
        set
        {
            Xml.AddElementVal(Name.Emphasis, value != Emphasis.None ? value.GetEnumName() : null);
            setProperties.Add(nameof(Emphasis));
        }
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

            setProperties.Add(nameof(StrikeThrough));
        }
    }

    /// <summary>
    /// The shade pattern applied to this paragraph
    /// </summary>
    public Shading Shading => new(Xml);

    // TODO: add TextBorder (bdr)

    /// <summary>
    /// Constructor
    /// </summary>
    public Formatting()
    {
        Xml = new XElement(Name.RunProperties);
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="element">Element containing run properties</param>
    internal Formatting(XElement element)
    {
        Xml = element;
    }

    /// <summary>
    /// Equals method to compare to another formatting object.
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Formatting? other)
    {
        return ReferenceEquals(this, other) 
               || (other != null && XNode.DeepEquals(Xml.Normalize(), other.Xml.Normalize()));
    }

    /// <summary>
    /// Determines equality for run properties
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Formatting);

    /// <summary>
    /// Returns hashcode for this formatting object
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();

    /// <summary>
    /// Merge in the given formatting into this formatting object. This will add/remove settings
    /// from the given formatting into this one.
    /// </summary>
    /// <param name="other">Formatting to merge in</param>
    public void Merge(Formatting other)
    {
        // First merge any changed properties.
        foreach (var propertyName in other.setProperties)
        {
            var propInfo = GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            if (propInfo == null)
                throw new InvalidOperationException($"Could not locate property setter for {propertyName} on {GetType().Name}");
            propInfo.SetValue(this, propInfo.GetValue(other));
        }

        // Now walk the XML graph and copy over any specifically set element.
        foreach (var el in other.Xml.Descendants())
        {
            if (Xml.Element(el.Name) == null)
                Xml.Add(el.Clone());
        }
    }
}