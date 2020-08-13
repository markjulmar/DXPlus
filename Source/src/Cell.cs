using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.Xml.Linq;

namespace DXPlus
{
    public class Cell : Container
    {
        internal Row row;

        internal Cell(Row row, DocX document, XElement xml)
            : base(document, xml)
        {
            this.row = row;
            this.packagePart = row.packagePart;
        }

        /// <summary>
        /// Gets or Sets the fill color of this Cell.
        /// </summary>
        public Color FillColor
        {
            get
            {
                XAttribute fill = Xml.Element(DocxNamespace.Main + "tcPr")?
                                    .Element(DocxNamespace.Main + "shd")?
                                    .Attribute(DocxNamespace.Main + "fill");

                if (fill != null)
                {
                    int argb = int.Parse(fill.Value.Replace("#", ""), NumberStyles.HexNumber);
                    return Color.FromArgb(argb);
                }

                return Color.Empty;
            }

            set
            {
                XElement tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                XElement shd = tcPr.GetOrCreateElement(DocxNamespace.Main + "shd");

                shd.SetAttributeValue(DocxNamespace.Main + "val", "clear");
                shd.SetAttributeValue(DocxNamespace.Main + "color", "auto");
                shd.SetAttributeValue(DocxNamespace.Main + "fill", value.ToHex());
            }
        }

        private double GetMargin(string name)
        {
            var w = Xml.Element(DocxNamespace.Main + "tcPr")?
                    .Element(DocxNamespace.Main + "tcMar")?
                    .Element(DocxNamespace.Main + name)?
                    .Attribute(DocxNamespace.Main + "w");

            if (w == null || !double.TryParse(w.Value, out double margin))
            {
                w?.Remove();
                return double.NaN;
            }

            // Convert margin to pixels - 15 units is equal to one pixel.
            return margin / 15;
        }

        private void SetMargin(string name, double value)
        {
            XElement tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
            XElement tcMar = tcPr.GetOrCreateElement(DocxNamespace.Main + "tcMar");
            XElement tcMarBottom = tcMar.GetOrCreateElement(DocxNamespace.Main + name);

            // The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
            tcMarBottom.SetAttributeValue(DocxNamespace.Main + "type", "dxa");
            // 15 "word units" is equal to one pixel.
            tcMarBottom.SetAttributeValue(DocxNamespace.Main + "w", value * 15);
        }


        /// <summary>
        /// BottomMargin in pixels.
        /// </summary>
        public double BottomMargin
        {
            get => GetMargin("bottom");
            set => SetMargin("bottom", value);
        }

        /// <summary>
        /// LeftMargin in pixels.
        /// </summary>
        public double LeftMargin
        {
            get => GetMargin("left");
            set => SetMargin("left", value);
        }

        /// <summary>
        /// RightMargin in pixels.
        /// </summary>
        public double RightMargin
        {
            get => GetMargin("right");
            set => SetMargin("right", value);
        }

        /// <summary>
        /// TopMargin in pixels.
        /// </summary>
        public double TopMargin
        {
            get => GetMargin("top");
            set => SetMargin("top", value);
        }

        public override ReadOnlyCollection<Paragraph> Paragraphs
        {
            get
            {
                var paragraphs = base.Paragraphs;
                foreach (Paragraph p in paragraphs)
                    p.packagePart = row.table.packagePart;

                return paragraphs;
            }
        }

        public Color Shading
        {
            get
            {
                var fill = Xml.Element(DocxNamespace.Main + "tcPr")?
                   .Element(DocxNamespace.Main + "shd")?
                   .Attribute(DocxNamespace.Main + "fill");

                return fill == null 
                    ? Color.White 
                    : ColorTranslator.FromHtml(string.Format("#{0}", fill.Value));
            }

            set
            {
                XElement tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                XElement shd = tcPr.GetOrCreateElement(DocxNamespace.Main + "shd");

                // The val attribute needs to be set to clear
                shd.SetAttributeValue(DocxNamespace.Main + "val", "clear");
                // The color attribute needs to be set to auto
                shd.SetAttributeValue(DocxNamespace.Main + "color", "auto");
                // The fill attribute needs to be set to the hex for this Color.
                shd.SetAttributeValue(DocxNamespace.Main + "fill", value.ToHex());
            }
        }

        public TextDirection TextDirection
        {
            get
            {
                XAttribute val = Xml.Element(DocxNamespace.Main + "tcPr")?
                   .Element(DocxNamespace.Main + "textDirection")?
                   .GetValAttr();

                if (val == null || !val.TryGetEnumValue<TextDirection>(out var result))
                {
                    val?.Remove();
                    return TextDirection.LeftToRightTopToBottom;
                }

                return result;
            }
            set
            {
                XElement tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                XElement textDirection = tcPr.GetOrCreateElement(DocxNamespace.Main + "textDirection");
                textDirection.SetVal(value.GetEnumName());
            }
        }

        /// <summary>
        /// Gets or Sets this Cells vertical alignment.
        /// </summary>
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                var val = Xml.Element(DocxNamespace.Main + "tcPr")?
                             .Element(DocxNamespace.Main + "vAlign")?
                             .GetValAttr();
                
                if (val == null || !val.TryGetEnumValue<VerticalAlignment>(out var result))
                {
                    val?.Remove();
                    return VerticalAlignment.Center;
                }
                
                return result;
            }

            set
            {
                XElement tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                XElement vAlign = tcPr.GetOrCreateElement(DocxNamespace.Main + "vAlign");
                vAlign.SetVal(value.GetEnumName());
            }
        }
        /// <summary>
        /// Width in pixels. // Added by Joel, refactored by Cathal
        /// </summary>
        public double Width
        {
            get
            {
                var value = Xml.Element(DocxNamespace.Main + "tcPr")?
                        .Element(DocxNamespace.Main + "tcW")?
                        .Attribute(DocxNamespace.Main + "w");

                if (value == null || !double.TryParse(value.Value, out double widthUnits))
                {
                    value?.Remove();
                    return double.NaN;
                }

                return widthUnits / 15;
            }

            set
            {
                XElement tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                XElement tcW = tcPr.GetOrCreateElement(DocxNamespace.Main + "tcW");

                if (value < 0)
                {
                    tcW.Remove();
                }
                else
                {
                    // The type attribute needs to be set to dxa which represents "twips" or twentieths of a point. In other words, 1/1440th of an inch.
                    tcW.SetAttributeValue(DocxNamespace.Main + "type", "dxa");
                    // 15 "word units" is equal to one pixel.
                    tcW.SetAttributeValue(DocxNamespace.Main + "w", value * 15);
                }
            }
        }
        /// <summary>
        /// Get a table cell border
        /// </summary>
        /// <param name="borderType">The table cell border to get</param>
        public Border GetBorder(TableCellBorderType borderType)
        {
            Border border = new Border();
            XElement tcBorder = Xml.Element(DocxNamespace.Main + "tcPr")?
                                    .Element(DocxNamespace.Main + "tcBorders")?
                                    .Element(DocxNamespace.Main + borderType.GetEnumName());
            if (tcBorder != null)
                border.GetDetails(tcBorder);

            return border;
        }

        /// <summary>
        /// Set the table cell border
        /// </summary>
        /// <param name="borderType">Table Cell border to set</param>
        /// <param name="border">Border object to set the table cell border</param>
        public void SetBorder(TableCellBorderType borderType, Border border)
        {
            XElement tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
            XElement tcBorders = tcPr.GetOrCreateElement(DocxNamespace.Main + "tcBorders");
            XElement tcBorderType = tcBorders.GetOrCreateElement(DocxNamespace.Main.NamespaceName + borderType.GetEnumName());
            
            // The val attribute is used for the style
            tcBorderType.SetVal(border.Style.GetEnumName());
            
            var size = border.Size switch
            {
                BorderSize.One => 2,
                BorderSize.Two => 4,
                BorderSize.Three => 6,
                BorderSize.Four => 8,
                BorderSize.Five => 12,
                BorderSize.Six => 18,
                BorderSize.Seven => 24,
                BorderSize.Eight => 36,
                BorderSize.Nine => 48,
                _ => 2,
            };

            // The sz attribute is used for the border size
            tcBorderType.SetAttributeValue(DocxNamespace.Main + "sz", size);

            // The space attribute is used for the cell spacing (probably '0')
            tcBorderType.SetAttributeValue(DocxNamespace.Main + "space", border.Space);

            // The color attribute is used for the border color
            tcBorderType.SetAttributeValue(DocxNamespace.Main + "color", border.Color.ToHex());
        }
    }
}