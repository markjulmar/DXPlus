using System;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// A single cell in a Word table. All content in a table is contained in a cell.
    /// A cell also has several properties affecting its size, appearance, and how the content it contains is formatted.
    /// </summary>
    public class Cell : BlockContainer
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="row">Owner Row</param>
        /// <param name="xml">XML representing this cell</param>
        internal Cell(Row row, XElement xml) : base(row.Document, xml)
        {
            Row = row;
            PackagePart = row.PackagePart;
        }

        /// <summary>
        /// Row owner
        /// </summary>
        public Row Row { get; }

        /// <summary>
        /// The table cell properties
        /// </summary>
        private XElement tcPr => Xml.GetOrAddElement(Namespace.Main + "tcPr");

        /// <summary>
        /// Gets or Sets the fill color of this Cell.
        /// </summary>
        public Color FillColor
        {
            get => tcPr.Element(Namespace.Main + "shd")?
                      .Attribute(Namespace.Main + "fill")
                      .ToColor() ?? Color.Empty;

            set
            {
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
            int.TryParse(tcPr.Element(Namespace.Main + "gridSpan")?
                .GetVal(), out int result)
                ? result
                : 1;

        /// <summary>
        /// Bottom margin in pixels.
        /// </summary>
        public double? BottomMargin
        {
            get => GetMargin(TableCellMarginType.Bottom.GetEnumName());
            set => SetMargin(TableCellMarginType.Bottom.GetEnumName(), value);
        }

        /// <summary>
        /// Left margin in pixels.
        /// </summary>
        public double? LeftMargin
        {
            get => GetMargin(TableCellMarginType.Left.GetEnumName());
            set => SetMargin(TableCellMarginType.Left.GetEnumName(), value);
        }

        /// <summary>
        /// Right margin in pixels.
        /// </summary>
        public double? RightMargin
        {
            get => GetMargin(TableCellMarginType.Right.GetEnumName());
            set => SetMargin(TableCellMarginType.Right.GetEnumName(), value);
        }

        /// <summary>
        /// Top margin in pixels.
        /// </summary>
        public double? TopMargin
        {
            get => GetMargin(TableCellMarginType.Top.GetEnumName());
            set => SetMargin(TableCellMarginType.Top.GetEnumName(), value);
        }

        /// <summary>
        /// Shading applied to the cell.
        /// </summary>
        public Color Shading
        {
            get
            {
                var fill = tcPr.Element(Namespace.Main + "shd")?
                    .Attribute(Namespace.Main + "fill");

                return fill == null
                    ? Color.White
                    : ColorTranslator.FromHtml($"#{fill.Value}");
            }

            set
            {
                var shd = tcPr.GetOrAddElement(Namespace.Main + "shd");

                // The val attribute needs to be set to clear
                shd.SetAttributeValue(Name.MainVal, "clear");
                // The color attribute needs to be set to auto
                shd.SetAttributeValue(Name.Color, "auto");
                // The fill attribute needs to be set to the hex for this Color.
                shd.SetAttributeValue(Namespace.Main + "fill", value.ToHex());
            }
        }

        /// <summary>
        /// Gets or sets all the text for a paragraph. This will replace any existing paragraph(s)
        /// tied to the table. The <seealso cref="Paragraph"/> property can also be used to manipulate
        /// the content of the cell.
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
                XAttribute val = tcPr.Element(Namespace.Main + "textDirection")?.GetValAttr();
                if (!val.TryGetEnumValue(out TextDirection result))
                {
                    val?.Remove();
                    return TextDirection.LeftToRightTopToBottom;
                }

                return result;
            }
            set
            {
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
                XAttribute val = tcPr.Element(Namespace.Main + "vAlign")?.GetValAttr();
                if (!val.TryGetEnumValue(out VerticalAlignment result))
                {
                    val?.Remove();
                    return VerticalAlignment.Center;
                }
                return result;
            }

            set
            {
                XElement vAlign = tcPr.GetOrAddElement(Namespace.Main + "vAlign");
                vAlign.SetAttributeValue(Name.MainVal, value.GetEnumName());
            }
        }

        /// <summary>
        /// The units this colum width is expressed in.
        /// </summary>
        public TableWidthUnit WidthUnit => Enum.TryParse<TableWidthUnit>(tcPr.Element(Namespace.Main + "tcW")?
                                                    .AttributeValue(Namespace.Main + "type"), ignoreCase: true, out var tbw) ? tbw : TableWidthUnit.Auto;

        /// <summary>
        /// Width in points
        /// </summary>
        public double? Width
        {
            get
            {
                XAttribute value = tcPr.Element(Namespace.Main + "tcW")?.Attribute(Namespace.Main + "w");
                if (value == null || !double.TryParse(value.Value, out double widthUnits))
                {
                    value?.Remove();
                    return null;
                }

                return widthUnits;
            }
        }

        /// <summary>
        /// Sets the column width
        /// </summary>
        /// <param name="unitType">Unit type</param>
        /// <param name="value">Width in dxa or % units</param>
        public void SetWidth(TableWidthUnit unitType, double? value)
        {
            XElement tcW = tcPr.GetOrAddElement(Namespace.Main + "tcW");
            if (unitType == TableWidthUnit.None || value == null || value < 0)
            {
                tcW.Remove();
                return;
            }

            tcW.SetAttributeValue(Namespace.Main + "type", unitType.GetEnumName());
            if (unitType == TableWidthUnit.Auto)
                value = 0;

            if (unitType == TableWidthUnit.Percentage)
                tcW.SetAttributeValue(Namespace.Main + "w", value.Value + "%");
            else
                tcW.SetAttributeValue(Namespace.Main + "w", value.Value);
        }

        /// <summary>
        /// Method to set all margins for the cell.
        /// </summary>
        /// <param name="margin"></param>
        public void SetMargins(double? margin)
        {
            LeftMargin = RightMargin = TopMargin = BottomMargin = margin;
        }

        /// <summary>
        /// Internal method to retrieve a specific margin by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Margin in dxa units</returns>
        private double? GetMargin(string name)
        {
            XAttribute w = tcPr.Element(Namespace.Main + "tcMar")?
                .Element(Namespace.Main + name)?
                .Attribute(Namespace.Main + "w");

            if (w == null || !double.TryParse(w.Value, out double margin))
            {
                w?.Remove();
                return null;
            }

            return margin;
        }

        /// <summary>
        /// Internal method to set a specific margin by name
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value in dxa units"></param>
        private void SetMargin(string name, double? value)
        {
            if (value != null)
            {
                XElement tcMar = tcPr.GetOrAddElement(Namespace.Main + "tcMar");
                XElement margin = tcMar.GetOrAddElement(Namespace.Main + name);

                margin.SetAttributeValue(Namespace.Main + "type", "dxa");
                margin.SetAttributeValue(Namespace.Main + "w", value);
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
        public Border GetBorder(TableCellBorderType borderType)
        {
            var tcBorder = tcPr.Element(Namespace.Main + "tcBorders")?
                .Element(Namespace.Main + borderType.GetEnumName());
            
            return tcBorder != null ? new Border(tcBorder) : null;
        }

        /// <summary>
        /// Set outside borders to the given style
        /// </summary>
        public void SetOutsideBorders(BorderStyle style, Color color, double? spacing = 1, double size = 2)
        {
            SetBorder(TableCellBorderType.Top, style, color, spacing, size);
            SetBorder(TableCellBorderType.Left, style, color, spacing, size);
            SetBorder(TableCellBorderType.Right, style, color, spacing, size);
            SetBorder(TableCellBorderType.Bottom, style, color, spacing, size);
        }

        /// <summary>
        /// Set the table cell border
        /// </summary>
        public void SetBorder(TableCellBorderType borderType, BorderStyle style, Color color, double? spacing = 1, double size = 2)
        {
            if (size is < 2 or > 96)
                throw new ArgumentOutOfRangeException(nameof(Size));
            if (!Enum.IsDefined(typeof(ParagraphBorderType), borderType))
                throw new InvalidEnumArgumentException(nameof(borderType), (int)borderType, typeof(ParagraphBorderType));

            tcPr.Element(Namespace.Main + "tcBorders")?
                .Element(Namespace.Main.NamespaceName + borderType.GetEnumName())?.Remove();

            if (borderType == TableCellBorderType.None)
                return;

            // Set the border style
            XElement tcBorders = tcPr.GetOrAddElement(Namespace.Main + "tcBorders");
            var borderXml = new XElement(Namespace.Main + borderType.GetEnumName(),
                new XAttribute(Name.MainVal, style.GetEnumName()),
                new XAttribute(Name.Size, size));
            if (color != Color.Empty)
                borderXml.Add(new XAttribute(Name.Color, color.ToHex()));
            if (spacing != null)
                borderXml.Add(new XAttribute(Namespace.Main + "space", spacing));

            tcBorders.Add(borderXml);
        }
    }
}