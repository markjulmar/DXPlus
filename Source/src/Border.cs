using System;
using System.Drawing;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a border of a table or table cell
    /// </summary>
    public class Border
    {
        /// <summary>
        /// Border style (dashed, line, etc.)
        /// </summary>
        public BorderStyle Style { get; set; }

        /// <summary>
        /// Border size
        /// </summary>
        public BorderSize Size { get; set; }

        /// <summary>
        /// Specifies the spacing offset that shall be used to place this border on the parent object
        /// </summary>
        public int SpacingOffset { get; set; }

        /// <summary>
        /// Border color
        /// </summary>
        public Color Color { get; set; }

        /// <summary>
        /// Default border constructor
        /// </summary>
        public Border() : this(BorderStyle.Single, BorderSize.One, 0, Color.Black)
        {
        }

        /// <summary>
        /// Parameterized border constructor
        /// </summary>
        /// <param name="style">Border style</param>
        /// <param name="size">Border size</param>
        /// <param name="spacingOffset">Border spacing</param>
        /// <param name="color">Border color</param>
        public Border(BorderStyle style, BorderSize size, int spacingOffset, Color color)
        {
            Style = style;
            Size = size;
            SpacingOffset = spacingOffset;
            Color = color;
        }

        /// <summary>
        /// Collects the details for the border from a tcBorders/[top|bottom|left|right] XML element
        /// <w:tcPr>
        ///   <w:tcBorders>
        ///     <w:top w:val="double" w:sz="24" w:spacingOffset="0" w:color="FF0000"/>
        ///     <w:left w:val="double" w:sz="24" w:spacingOffset="0" w:color="FF0000"/>
        ///     <w:bottom w:val="double" w:sz="24" w:spacingOffset="0" w:color="FF0000"/>
        ///     <w:right w:val="double" w:sz="24" w:spacingOffset="0" w:color="FF0000"/>
        ///   </w:tcBorders>
        /// </w:tcPr>        
        /// </summary>
        /// <param name="borderDetails"></param>
        internal void GetDetails(XElement borderDetails)
        {
            if (borderDetails == null) 
                return;
            
            // The val attribute is the border style
            var val = borderDetails.Attribute(Name.MainVal);
            if (val != null)
            {
                if (val.TryGetEnumValue<BorderStyle>(out var bs))
                {
                    Style = bs;
                }
                else
                {
                    val.Remove();
                }
            }

            // The sz attribute is used for the border size
            var sz = borderDetails.Attribute(Name.Size);
            if (sz != null)
            {
                if (int.TryParse(sz.Value, out int result))
                {
                    Size = result switch
                    {
                        2 => BorderSize.One,
                        4 => BorderSize.Two,
                        6 => BorderSize.Three,
                        8 => BorderSize.Four,
                        12 => BorderSize.Five,
                        18 => BorderSize.Six,
                        24 => BorderSize.Seven,
                        36 => BorderSize.Eight,
                        48 => BorderSize.Nine,
                        _ => BorderSize.One,
                    };
                }
                else
                {
                    sz.Remove();
                }
            }

            // The space attribute is used for the border spacing
            var space = borderDetails.Attribute(Namespace.Main + "space");
            if (space != null)
            {
                if (int.TryParse(space.Value, out int result))
                {
                    SpacingOffset = result;
                }
                else
                {
                    space.Remove();
                }
            }

            // The color attribute is used for the border color
            var color = borderDetails.Attribute(Name.Color);
            if (color != null)
            {
                try
                {
                    Color = ColorTranslator.FromHtml($"#{color.Value}");
                }
                catch
                {
                    color.Remove();
                }
            }
        }

    }
}