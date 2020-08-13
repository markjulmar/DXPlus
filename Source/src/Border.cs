using System;
using System.Drawing;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a border of a table or table cell
    /// </summary>
    public class Border
    {
        public BorderStyle Style { get; set; }
        public BorderSize Size { get; set; }
        public int Space { get; set; }
        public Color Color { get; set; }
        
        public Border() : this(BorderStyle.Single, BorderSize.One, 0, Color.Black)
        {
        }

        public Border(BorderStyle style, BorderSize size, int space, Color color)
        {
            Style = style;
            Size = size;
            Space = space;
            Color = color;
        }

		internal void GetDetails(XElement borderDetails)
		{
			if (borderDetails != null)
			{
				// The val attribute is the border style
				XAttribute val = borderDetails.Attribute(DocxNamespace.Main + "val");
				if (val != null)
				{
					if (Enum.TryParse<BorderStyle>(val.Value, ignoreCase: true, out var result))
					{
						this.Style = result;
					}
					else
					{
						val.Remove();
					}
				}

				// The sz attribute is used for the border size
				XAttribute sz = borderDetails.Attribute(DocxNamespace.Main + "sz");
				if (sz != null)
				{
					if (int.TryParse(sz.Value, out int result))
					{
						this.Size = result switch
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

				// The space attribute is used for the border spacing (probably '0')
				XAttribute space = borderDetails.Attribute(DocxNamespace.Main + "space");
				if (space != null)
				{
					if (int.TryParse(space.Value, out int result))
					{
						this.Space = result;
					}
					else
					{
						space.Remove();
					}
				}

				// The color attribute is used for the border color
				XAttribute color = borderDetails.Attribute(DocxNamespace.Main + "color");
				if (color != null)
				{
					try
					{
						this.Color = ColorTranslator.FromHtml(string.Format("#{0}", color.Value));
					}
					catch
					{
						color.Remove();
					}
				}
			}
		}

	}
}