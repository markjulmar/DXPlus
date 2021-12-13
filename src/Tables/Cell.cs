using DXPlus.Helpers;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// A single cell in a Word table
    /// </summary>
    public class Cell : BlockContainer
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="row">Owner Row</param>
        /// <param name="xml">XML representing this cell</param>
        internal Cell(Row row, XElement xml)
            : base(row.Document, xml)
        {
            Row = row;
        }

        /// <summary>
        /// Row owner
        /// </summary>
        public Row Row { get; set; }

        /// <summary>
        ///     Gets or Sets the fill color of this Cell.
        /// </summary>
        public Color FillColor
        {
            get => Xml.Element(Namespace.Main + "tcPr")?
                      .Element(Namespace.Main + "shd")?
                      .Attribute(Namespace.Main + "fill")
                      .ToColor() ?? Color.Empty;

            set
            {
                XElement tcPr = Xml.GetOrAddElement(Namespace.Main + "tcPr");
                XElement shd = tcPr.GetOrAddElement(Namespace.Main + "shd");

                shd.SetAttributeValue(Name.MainVal, "clear");
                shd.SetAttributeValue(Name.Color, "auto");
                shd.SetAttributeValue(Namespace.Main + "fill", value.ToHex());
            }
        }

        /// <summary>
        /// Get the applied gridSpan based on cell merges.
        /// </summary>
        public int GridSpan =>
            int.TryParse(Xml.Element(Namespace.Main + "tcPr")?
                .Element(Namespace.Main + "gridSpan")?
                .GetVal(), out int result)
                ? result
                : 1;

        /// <summary>
        /// BottomMargin in pixels.
        /// </summary>
        public double? BottomMargin
        {
            get => GetMargin(TableCellMarginType.Bottom.GetEnumName());
            set => SetMargin(TableCellMarginType.Bottom.GetEnumName(), value);
        }

        /// <summary>
        /// LeftMargin in pixels.
        /// </summary>
        public double? LeftMargin
        {
            get => GetMargin(TableCellMarginType.Left.GetEnumName());
            set => SetMargin(TableCellMarginType.Left.GetEnumName(), value);
        }

        /// <summary>
        ///     RightMargin in pixels.
        /// </summary>
        public double? RightMargin
        {
            get => GetMargin(TableCellMarginType.Right.GetEnumName());
            set => SetMargin(TableCellMarginType.Right.GetEnumName(), value);
        }

        /// <summary>
        ///     TopMargin in pixels.
        /// </summary>
        public double? TopMargin
        {
            get => GetMargin(TableCellMarginType.Top.GetEnumName());
            set => SetMargin(TableCellMarginType.Top.GetEnumName(), value);
        }

        public Color Shading
        {
            get
            {
                XAttribute fill = Xml.Element(Namespace.Main + "tcPr")?
                    .Element(Namespace.Main + "shd")?
                    .Attribute(Namespace.Main + "fill");

                return fill == null
                    ? Color.White
                    : ColorTranslator.FromHtml($"#{fill.Value}");
            }

            set
            {
                XElement tcPr = Xml.GetOrAddElement(Namespace.Main + "tcPr");
                XElement shd = tcPr.GetOrAddElement(Namespace.Main + "shd");

                // The val attribute needs to be set to clear
                shd.SetAttributeValue(Name.MainVal, "clear");
                // The color attribute needs to be set to auto
                shd.SetAttributeValue(Name.Color, "auto");
                // The fill attribute needs to be set to the hex for this Color.
                shd.SetAttributeValue(Namespace.Main + "fill", value.ToHex());
            }
        }

        /// <summary>
        /// Gets or sets all the text for a paragraph
        /// </summary>
        public string Text
        {
            get => string.Join('\n', Paragraphs.Select(p => p.Text).Where(p => !string.IsNullOrEmpty(p)));
            set
            {
                string val = value ?? string.Empty;
                switch (Paragraphs.Count())
                {
                    case 0:
                        this.AddParagraph(val);
                        break;

                    case 1:
                        Paragraphs.First().SetText(val);
                        break;

                    default:
                        Xml.Elements(Name.Paragraph).Remove();
                        this.AddParagraph(val);
                        break;
                }
            }
        }

        /// <summary>
        /// Set the text direction for the table cell
        /// </summary>
        public TextDirection TextDirection
        {
            get
            {
                XAttribute val = Xml.Element(Namespace.Main + "tcPr")?
                    .Element(Namespace.Main + "textDirection")?
                    .GetValAttr();

                if (!val.TryGetEnumValue(out TextDirection result))
                {
                    val?.Remove();
                    return TextDirection.LeftToRightTopToBottom;
                }

                return result;
            }
            set
            {
                XElement tcPr = Xml.GetOrAddElement(Namespace.Main + "tcPr");
                XElement textDirection = tcPr.GetOrAddElement(Namespace.Main + "textDirection");
                textDirection.SetAttributeValue(Name.MainVal, value.GetEnumName());
            }
        }

        /// <summary>
        /// Gets or Sets the vertical alignment.
        /// </summary>
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                XAttribute val = Xml.Element(Namespace.Main + "tcPr")?
                    .Element(Namespace.Main + "vAlign")?
                    .GetValAttr();

                if (!val.TryGetEnumValue(out VerticalAlignment result))
                {
                    val?.Remove();
                    return VerticalAlignment.Center;
                }
                return result;
            }

            set
            {
                XElement tcPr = Xml.GetOrAddElement(Namespace.Main + "tcPr");
                XElement vAlign = tcPr.GetOrAddElement(Namespace.Main + "vAlign");
                vAlign.SetAttributeValue(Name.MainVal, value.GetEnumName());
            }
        }

        /// <summary>
        /// Width in pixels
        /// </summary>
        public double? Width
        {
            get
            {
                XAttribute value = Xml.Element(Namespace.Main + "tcPr")?
                    .Element(Namespace.Main + "tcW")?
                    .Attribute(Namespace.Main + "w");

                if (value == null || !double.TryParse(value.Value, out double widthUnits))
                {
                    value?.Remove();
                    return null;
                }

                return widthUnits / TableHelpers.UnitConversion;
            }

            set
            {
                XElement tcPr = Xml.GetOrAddElement(Namespace.Main + "tcPr");
                XElement tcW = tcPr.GetOrAddElement(Namespace.Main + "tcW");

                if (value == null || value < 0)
                {
                    tcW.Remove();
                }
                else
                {
                    tcW.SetAttributeValue(Namespace.Main + "type", "dxa"); // Widths in 20th/pt.
                    tcW.SetAttributeValue(Namespace.Main + "w", value * TableHelpers.UnitConversion);
                }
            }
        }

        /// <summary>
        /// Method to set all margins for the cell.
        /// </summary>
        /// <param name="margin"></param>
        public void SetMargins(double? margin)
        {
            LeftMargin = RightMargin = TopMargin = BottomMargin = margin;
        }

        private double? GetMargin(string name)
        {
            XAttribute w = Xml.Element(Namespace.Main + "tcPr")?
                .Element(Namespace.Main + "tcMar")?
                .Element(Namespace.Main + name)?
                .Attribute(Namespace.Main + "w");

            if (w == null || !double.TryParse(w.Value, out double margin))
            {
                w?.Remove();
                return null;
            }

            // Convert margin to pixels.
            return margin / TableHelpers.UnitConversion;
        }

        private void SetMargin(string name, double? value)
        {
            XElement tcPr = Xml.GetOrAddElement(Namespace.Main + "tcPr");

            if (value != null)
            {
                XElement tcMar = tcPr.GetOrAddElement(Namespace.Main + "tcMar");
                XElement margin = tcMar.GetOrAddElement(Namespace.Main + name);

                margin.SetAttributeValue(Namespace.Main + "type", "dxa");
                margin.SetAttributeValue(Namespace.Main + "w", value * TableHelpers.UnitConversion);
            }
            else
            {
                tcPr.Element("tcMar")?.Element(Namespace.Main + name)?.Remove();
            }
        }

        /// <summary>
        /// Get a table cell border
        /// </summary>
        /// <param name="borderType">The table cell border to get</param>
        public TableBorder GetBorder(TableCellBorderType borderType)
        {
            TableBorder border = new TableBorder();
            XElement tcBorder = Xml.Element(Namespace.Main + "tcPr")?
                .Element(Namespace.Main + "tcBorders")?
                .Element(Namespace.Main + borderType.GetEnumName());
            if (tcBorder != null)
            {
                border.GetDetails(tcBorder);
            }

            return border;
        }

        /// <summary>
        /// Set the table cell border
        /// </summary>
        /// <param name="borderType">Table Cell border to set</param>
        /// <param name="tableBorder">Border object to set the table cell border</param>
        public void SetBorder(TableCellBorderType borderType, TableBorder tableBorder)
        {
            XElement tcPr = Xml.GetOrAddElement(Namespace.Main + "tcPr");
            XElement tcBorders = tcPr.GetOrAddElement(Namespace.Main + "tcBorders");
            XElement tcBorderType =
                tcBorders.GetOrAddElement(Namespace.Main.NamespaceName + borderType.GetEnumName());

            // The val attribute is used for the style
            tcBorderType.SetAttributeValue(Name.MainVal, tableBorder.Style.GetEnumName());

            int size = tableBorder.Size switch
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
                _ => 2
            };

            // The sz attribute is used for the border size
            tcBorderType.SetAttributeValue(Name.Size, size);

            // The space attribute is used for the cell spacing (probably '0')
            tcBorderType.SetAttributeValue(Namespace.Main + "space", tableBorder.SpacingOffset);

            // The color attribute is used for the border color
            tcBorderType.SetAttributeValue(Name.Color, tableBorder.Color.ToHex());
        }
    }
}