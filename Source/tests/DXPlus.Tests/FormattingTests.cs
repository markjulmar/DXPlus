﻿using System;
using System.Drawing;
using System.Globalization;
using System.Xml.XPath;
using Xunit;

namespace DXPlus.Tests
{
    public class FormattingTests
    {
        [Fact]
        public void BoldAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.False(rPr.Bold);

            rPr.Bold = true;
            Assert.True(rPr.Bold);
            Assert.Single(rPr.Xml.Elements(Name.Bold));
            rPr.Bold = true;
            Assert.Single(rPr.Xml.Elements(Name.Bold));
            rPr.Bold = false;
            Assert.False(rPr.Bold);
            Assert.Empty(rPr.Xml.Elements());
        }

        [Fact]
        public void ItalicAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.False(rPr.Italic);

            rPr.Italic = true;
            Assert.True(rPr.Italic);
            Assert.Single(rPr.Xml.Elements(Name.Italic));
            rPr.Italic = true;
            Assert.Single(rPr.Xml.Elements(Name.Italic));
            rPr.Italic = false;
            Assert.False(rPr.Italic);
            Assert.Empty(rPr.Xml.Elements());
        }

        [Fact]
        public void CapsStyleAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Equal(CapsStyle.None, rPr.CapsStyle);

            rPr.CapsStyle = CapsStyle.Caps;
            Assert.Equal(CapsStyle.Caps, rPr.CapsStyle);

            Assert.NotNull(rPr.Xml.RemoveNamespaces().XPathSelectElement("caps"));
            Assert.Null(rPr.Xml.RemoveNamespaces().XPathSelectElement("smallCaps"));

            rPr.CapsStyle = CapsStyle.SmallCaps;
            Assert.Equal(CapsStyle.SmallCaps, rPr.CapsStyle);

            Assert.NotNull(rPr.Xml.RemoveNamespaces().XPathSelectElement("smallCaps"));
            Assert.Null(rPr.Xml.RemoveNamespaces().XPathSelectElement("caps"));

            rPr.CapsStyle = CapsStyle.None;
            Assert.Equal(CapsStyle.None, rPr.CapsStyle);

            Assert.Null(rPr.Xml.RemoveNamespaces().XPathSelectElement("smallCaps"));
            Assert.Null(rPr.Xml.RemoveNamespaces().XPathSelectElement("caps"));
        }

        [Fact]
        public void ColorAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Equal(Color.Empty, rPr.Color);

            rPr.Color = Color.Red;
            Assert.NotStrictEqual(Color.Red, rPr.Color);
            Assert.NotNull(rPr.Xml.RemoveNamespaces().XPathSelectElement("color[@val='FF0000']"));

            rPr.Color = Color.Empty;
            Assert.Equal(Color.Empty, rPr.Color);
            Assert.Null(rPr.Xml.RemoveNamespaces().XPathSelectElement("color"));
        }

        [Fact]
        public void CultureAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Null(rPr.Culture);

            var spanish = CultureInfo.GetCultureInfo("es-BR");

            rPr.Culture = spanish;
            Assert.Equal(spanish, rPr.Culture);
            Assert.NotNull(rPr.Xml.RemoveNamespaces().XPathSelectElement("lang[@val='es-BR']"));

            rPr.Culture = null;
            Assert.Null(rPr.Culture);
            Assert.Null(rPr.Xml.RemoveNamespaces().XPathSelectElement("lang"));
        }

        [Fact]
        public void FontAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Null(rPr.Font);

            var ff = new FontFamily("Times New Roman");

            rPr.Font = ff;
            Assert.Equal(ff.Name, rPr.Font.Name);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("rFonts[@ascii='Times New Roman']"));

            ff = new FontFamily("Wingdings");
            rPr.Font = ff;
            Assert.Equal(ff.Name, rPr.Font.Name);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("rFonts"));

            rPr.Font = null;
            Assert.Null(rPr.Font);
            Assert.Null(rPr.Xml.RemoveNamespaces().XPathSelectElement("rFonts"));
        }

        [Fact]
        public void FontSizeAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Null(rPr.FontSize);

            rPr.FontSize = 32;
            Assert.Equal(32, rPr.FontSize);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("sz"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("szCs"));
            Assert.Equal("64", rPr.Xml.RemoveNamespaces().XPathSelectElement("sz").Attribute("val")?.Value);

            rPr.FontSize = 22;
            Assert.Equal(22, rPr.FontSize);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("sz"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("szCs"));
            Assert.Equal("44", rPr.Xml.RemoveNamespaces().XPathSelectElement("sz").Attribute("val")?.Value);

            rPr.FontSize = null;
            Assert.Null(rPr.FontSize);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("sz"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("szCs"));
        }

        [Fact]
        public void HideAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.False(rPr.IsHidden);

            rPr.IsHidden = true;
            Assert.True(rPr.IsHidden);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("vanish"));

            // Make sure we don't dup the tag
            rPr.IsHidden = true;
            Assert.True(rPr.IsHidden);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("vanish"));

            rPr.IsHidden = false;
            Assert.False(rPr.IsHidden);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("vanish"));
        }

        [Fact]
        public void HighlightAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Equal(Highlight.None, rPr.Highlight);

            rPr.Highlight = Highlight.Yellow;
            Assert.Equal(Highlight.Yellow, rPr.Highlight);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("highlight"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("highlight[@val='yellow']"));

            // Make sure we don't dup the tag
            rPr.Highlight = Highlight.Green;
            Assert.Equal(Highlight.Green, rPr.Highlight);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("highlight"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("highlight[@val='green']"));

            rPr.Highlight = Highlight.None;
            Assert.Equal(Highlight.None, rPr.Highlight);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("highlight"));
        }

        [Fact]
        public void KerningAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Null(rPr.Kerning);

            rPr.Kerning = 10;
            Assert.Equal(10, rPr.Kerning);

            rPr.Kerning = null;
            Assert.Null(rPr.Kerning);
        }

        [Fact]
        public void EffectAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Equal(Effect.None, rPr.Effect);

            rPr.Effect = Effect.Emboss;
            Assert.Equal(Effect.Emboss, rPr.Effect);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("emboss"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("shadow"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outline"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outlineShadow"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("imprint"));

            // Make sure we don't dup the tag
            rPr.Effect = Effect.Shadow;
            Assert.Equal(Effect.Shadow, rPr.Effect);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("emboss"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("shadow"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outline"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outlineShadow"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("imprint"));

            rPr.Effect = Effect.None;
            Assert.Equal(Effect.None, rPr.Effect);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("emboss"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("shadow"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outline"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outlineShadow"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("imprint"));

            rPr.Effect = Effect.OutlineShadow;
            Assert.Equal(Effect.OutlineShadow, rPr.Effect);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("emboss"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("shadow"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("outline"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outlineShadow"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("imprint"));

            rPr.Effect = Effect.Engrave;
            Assert.Equal(Effect.Engrave, rPr.Effect);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("emboss"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("shadow"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outline"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("outlineShadow"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("imprint"));
        }

        [Fact]
        public void ExpansionAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Null(rPr.ExpansionScale);

            rPr.ExpansionScale = 200;
            Assert.Equal(200, rPr.ExpansionScale);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("w"));

            rPr.ExpansionScale = null;
            Assert.Null(rPr.ExpansionScale);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("w"));

            Assert.Throws<ArgumentOutOfRangeException>(() => rPr.ExpansionScale = -1);
            Assert.Throws<ArgumentOutOfRangeException>(() => rPr.ExpansionScale = 601);
        }

        [Fact]
        public void PositionAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Null(rPr.Position);

            rPr.Position = 50;
            Assert.Equal(50, rPr.Position);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("position"));
            Assert.Equal("100", rPr.Xml.RemoveNamespaces().XPathSelectElement("position").Attribute("val")?.Value);

            rPr.Position = null;
            Assert.Null(rPr.Position);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("position"));

            rPr.Position = 10.23214452452;
            Assert.Equal(10.23, rPr.Position);

            rPr.Position = -12;
            Assert.Equal(-12, rPr.Position);
        }

        [Fact]
        public void SuperscriptAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.False(rPr.Superscript);
            Assert.False(rPr.Subscript);

            rPr.Superscript = true;
            Assert.True(rPr.Superscript);
            Assert.False(rPr.Subscript);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("vertAlign"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("vertAlign[@val='superscript']"));

            rPr.Superscript = false;
            Assert.False(rPr.Superscript);
            Assert.False(rPr.Subscript);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("vertAlign"));
        }

        [Fact]
        public void SuperscriptAffectSubscript()
        {
            var rPr = new Formatting();
            Assert.False(rPr.Superscript);
            Assert.False(rPr.Subscript);

            rPr.Superscript = true;
            Assert.True(rPr.Superscript);
            Assert.False(rPr.Subscript);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("vertAlign"));

            rPr.Subscript = true;
            Assert.False(rPr.Superscript);
            Assert.True(rPr.Subscript);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("vertAlign"));

            rPr.Superscript = false;
            Assert.False(rPr.Superscript);
            Assert.True(rPr.Subscript);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("vertAlign"));

            rPr.Subscript = false;
            Assert.False(rPr.Superscript);
            Assert.False(rPr.Subscript);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("vertAlign"));
        }

        [Fact]
        public void UnderlineAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Equal(UnderlineStyle.None, rPr.UnderlineStyle);

            rPr.UnderlineStyle = UnderlineStyle.SingleLine;
            Assert.Equal(UnderlineStyle.SingleLine, rPr.UnderlineStyle);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("u"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("u[@val='single']"));

            rPr.UnderlineStyle = UnderlineStyle.DoubleLine;
            Assert.Equal(UnderlineStyle.DoubleLine, rPr.UnderlineStyle);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("u"));

            rPr.UnderlineStyle = UnderlineStyle.None;
            Assert.Equal(UnderlineStyle.None, rPr.UnderlineStyle);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("u"));
        }

        [Fact]
        public void EmphasisAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Equal(Emphasis.None, rPr.Emphasis);

            rPr.Emphasis = Emphasis.Dot;
            Assert.Equal(Emphasis.Dot, rPr.Emphasis);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("em"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("em[@val='dot']"));

            rPr.Emphasis = Emphasis.Comma;
            Assert.Equal(Emphasis.Comma, rPr.Emphasis);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("em"));

            rPr.Emphasis = Emphasis.None;
            Assert.Equal(Emphasis.None, rPr.Emphasis);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("em"));
        }

        [Fact]
        public void StrikeThroughAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Equal(Strikethrough.None, rPr.StrikeThrough);

            rPr.StrikeThrough = Strikethrough.Strike;
            Assert.Equal(Strikethrough.Strike, rPr.StrikeThrough);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("strike"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("dstrike"));

            rPr.StrikeThrough = Strikethrough.DoubleStrike;
            Assert.Equal(Strikethrough.DoubleStrike, rPr.StrikeThrough);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("strike"));
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("dstrike"));

            rPr.StrikeThrough = Strikethrough.None;
            Assert.Equal(Strikethrough.None, rPr.StrikeThrough);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("strike"));
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("dstrike"));
        }

        [Fact]
        public void NoProofAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.False(rPr.NoProof);

            rPr.NoProof = true;
            Assert.True(rPr.NoProof);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("noProof"));

            // Make sure we don't dup the tag
            rPr.NoProof = true;
            Assert.True(rPr.NoProof);
            Assert.Single(rPr.Xml.RemoveNamespaces().XPathSelectElements("noProof"));

            rPr.NoProof = false;
            Assert.False(rPr.NoProof);
            Assert.Empty(rPr.Xml.RemoveNamespaces().XPathSelectElements("noProof"));
        }

        [Fact]
        public void UnderlineColorAddsRemovesElement()
        {
            var rPr = new Formatting();
            Assert.Equal(Color.Empty, rPr.UnderlineColor);

            rPr.UnderlineColor = Color.Red;
            Assert.NotStrictEqual(Color.Red, rPr.UnderlineColor);
            Assert.Equal(UnderlineStyle.SingleLine, rPr.UnderlineStyle);
            Assert.NotNull(rPr.Xml.RemoveNamespaces().XPathSelectElement("u[@color='FF0000']"));

            rPr.UnderlineColor = Color.Empty;
            Assert.Equal(Color.Empty, rPr.UnderlineColor);
            Assert.Null(rPr.Xml.RemoveNamespaces().XPathSelectElement("u[@color='FF0000']"));
        }

        [Fact]
        public void NewFormattingIsEqual()
        {
            var f1 = new Formatting();
            var f2 = new Formatting();
            Assert.True(f1.Equals(f2));
        }

        [Fact]
        public void BasicFormattingIsEqual()
        {
            var f1 = new Formatting { Bold = true, IsHidden =  true };
            var f2 = new Formatting { Bold = true, IsHidden = true };
            Assert.True(f1.Equals(f2));
        }

        [Fact]
        public void DifferentFormattingIsNotEqual()
        {
            var f1 = new Formatting { Bold = true, IsHidden = true, Italic = true };
            var f2 = new Formatting { Bold = true, IsHidden = true };
            Assert.False(f1.Equals(f2));
        }
    }
}