using System;
using System.Drawing;
using System.Globalization;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// This represents a border borderEdgeType along a side of an element.
    /// </summary>
    public class Border
    {
        private readonly XElement Xml;

        /// <summary>
        /// Returns the color for this border borderEdgeType.
        /// </summary>
        public Color Color
        {
            get => Xml.Attribute(Name.Color)?.ToColor() ?? Color.Empty;
            set => Xml.SetAttributeValue(Name.Color, value == Color.Empty ? null : value.ToHex());
        }

        /// <summary>
        /// Specifies whether the border should be modified to create the appearance of a shadow. For right and bottom borders, this is done
        /// by duplicating the border below and right of the normal location. For the right and top borders, this is done by moving the
        /// border down and to the right of the original location.
        /// </summary>
        public bool Shadow
        {
            get => bool.Parse(Xml.AttributeValue(Name.Shadow, "false"));
            set => Xml.SetAttributeValue(Name.Shadow, value);
        }

        /// <summary>
        /// Specifies the spacing offset. Values are specified in points (1/72nd of an inch).
        /// </summary>
        public double? Spacing
        {
            // Represented in 20ths of a pt.
            get => double.TryParse(Xml.AttributeValue(Namespace.Main + "space"), out var result) ? Math.Round(result / 20.0, 1) : null;
            set => Xml.AddElementVal(Namespace.Main + "space", value == null ? null : Math.Round(value.Value * 20.0, 1).ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Specifies the width of the border. Paragraph borders are line borders, the width is specified in eighths of a point
        /// with a minimum value of two (1/4 of a point) and a maximum value of 96 (twelve points).
        /// </summary>
        public double Size
        {
            get => double.Parse(Xml.AttributeValue(Name.Size));
            set
            {
                if (value is < 2 or > 96)
                    throw new ArgumentOutOfRangeException(nameof(Size));
                Xml.SetAttributeValue(Name.Size, value);
            }
        }

        /// <summary>
        /// Specifies the style of the border. Paragraph borders can be only line borders.
        /// </summary>
        public BorderStyle Style
        {
            get => Enum.TryParse<BorderStyle>(Xml.AttributeValue(Name.MainVal, "None"), out var bd) ? bd : BorderStyle.None;
            set => Xml.SetAttributeValue(Name.MainVal, value.GetEnumName());
        }

        /// <summary>
        /// Border edge for an existing border edge.
        /// </summary>
        /// <param name="xml"></param>
        internal Border(XElement xml)
        {
            Xml = xml;
        }
    }
}