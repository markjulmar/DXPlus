using System;
using System.Drawing;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// A single cell in a Word table
    /// </summary>
    public class Cell : Container
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
            PackagePart = row.PackagePart;
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
            get => Xml.Element(DocxNamespace.Main + "tcPr")?
                      .Element(DocxNamespace.Main + "shd")?
                      .Attribute(DocxNamespace.Main + "fill")
                      .ToColor() ?? Color.Empty;

            set
            {
                var tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                var shd = tcPr.GetOrCreateElement(DocxNamespace.Main + "shd");

                shd.SetAttributeValue(DocxNamespace.Main + "val", "clear");
                shd.SetAttributeValue(DocxNamespace.Main + "color", "auto");
                shd.SetAttributeValue(DocxNamespace.Main + "fill", value.ToHex());
            }
        }

        /// <summary>
        /// Get the applied gridSpan based on cell merges.
        /// </summary>
        public int GridSpan =>
            int.TryParse(Xml.Element(DocxNamespace.Main + "tcPr")?
                .Element(DocxNamespace.Main + "gridSpan")?
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
                var fill = Xml.Element(DocxNamespace.Main + "tcPr")?
                    .Element(DocxNamespace.Main + "shd")?
                    .Attribute(DocxNamespace.Main + "fill");

                return fill == null
                    ? Color.White
                    : ColorTranslator.FromHtml($"#{fill.Value}");
            }

            set
            {
                var tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                var shd = tcPr.GetOrCreateElement(DocxNamespace.Main + "shd");

                // The val attribute needs to be set to clear
                shd.SetAttributeValue(DocxNamespace.Main + "val", "clear");
                // The color attribute needs to be set to auto
                shd.SetAttributeValue(DocxNamespace.Main + "color", "auto");
                // The fill attribute needs to be set to the hex for this Color.
                shd.SetAttributeValue(DocxNamespace.Main + "fill", value.ToHex());
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
                string val = value ?? "";
                switch (Paragraphs.Count)
                {
                    case 0: InsertParagraph(val);
                        break;
                    case 1: Paragraphs[0].SetText(val);
                        break;
                    default:
                        Xml.Elements(DocxNamespace.Main + "p").Remove();
                        InsertParagraph(val);
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
                var val = Xml.Element(DocxNamespace.Main + "tcPr")?
                    .Element(DocxNamespace.Main + "textDirection")?
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
                var tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                var textDirection = tcPr.GetOrCreateElement(DocxNamespace.Main + "textDirection");
                textDirection.SetAttributeValue(DocxNamespace.Main + "val", value.GetEnumName());
            }
        }

        /// <summary>
        /// Gets or Sets the vertical alignment.
        /// </summary>
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                var val = Xml.Element(DocxNamespace.Main + "tcPr")?
                    .Element(DocxNamespace.Main + "vAlign")?
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
                var tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                var vAlign = tcPr.GetOrCreateElement(DocxNamespace.Main + "vAlign");
                vAlign.SetAttributeValue(DocxNamespace.Main + "val", value.GetEnumName());
            }
        }

        /// <summary>
        /// Width in pixels
        /// </summary>
        public double? Width
        {
            get
            {
                var value = Xml.Element(DocxNamespace.Main + "tcPr")?
                    .Element(DocxNamespace.Main + "tcW")?
                    .Attribute(DocxNamespace.Main + "w");

                if (value == null || !double.TryParse(value.Value, out var widthUnits))
                {
                    value?.Remove();
                    return null;
                }

                return widthUnits / TableHelpers.UnitConversion;
            }

            set
            {
                var tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
                var tcW = tcPr.GetOrCreateElement(DocxNamespace.Main + "tcW");

                if (value < 0)
                {
                    tcW.Remove();
                }
                else
                {
                    tcW.SetAttributeValue(DocxNamespace.Main + "type", "dxa"); // Widths in 20th/pt.
                    tcW.SetAttributeValue(DocxNamespace.Main + "w", value * TableHelpers.UnitConversion);
                }
            }
        }

        private double? GetMargin(string name)
        {
            var w = Xml.Element(DocxNamespace.Main + "tcPr")?
                .Element(DocxNamespace.Main + "tcMar")?
                .Element(DocxNamespace.Main + name)?
                .Attribute(DocxNamespace.Main + "w");

            if (w == null || !double.TryParse(w.Value, out var margin))
            {
                w?.Remove();
                return null;
            }

            // Convert margin to pixels.
            return margin / TableHelpers.UnitConversion;
        }

        private void SetMargin(string name, double? value)
        {
            var tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");

            if (value != null)
            {
                var tcMar = tcPr.GetOrCreateElement(DocxNamespace.Main + "tcMar");
                var margin = tcMar.GetOrCreateElement(DocxNamespace.Main + name);

                margin.SetAttributeValue(DocxNamespace.Main + "type", "dxa");
                margin.SetAttributeValue(DocxNamespace.Main + "w", value * TableHelpers.UnitConversion);
            }
            else
            {
                tcPr.Element("tcMar")?.Element(DocxNamespace.Main + name)?.Remove();
            }
        }

        /// <summary>
        /// Get a table cell border
        /// </summary>
        /// <param name="borderType">The table cell border to get</param>
        public Border GetBorder(TableCellBorderType borderType)
        {
            var border = new Border();
            var tcBorder = Xml.Element(DocxNamespace.Main + "tcPr")?
                .Element(DocxNamespace.Main + "tcBorders")?
                .Element(DocxNamespace.Main + borderType.GetEnumName());
            if (tcBorder != null) border.GetDetails(tcBorder);

            return border;
        }

        /// <summary>
        /// Set the table cell border
        /// </summary>
        /// <param name="borderType">Table Cell border to set</param>
        /// <param name="border">Border object to set the table cell border</param>
        public void SetBorder(TableCellBorderType borderType, Border border)
        {
            var tcPr = Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
            var tcBorders = tcPr.GetOrCreateElement(DocxNamespace.Main + "tcBorders");
            var tcBorderType =
                tcBorders.GetOrCreateElement(DocxNamespace.Main.NamespaceName + borderType.GetEnumName());

            // The val attribute is used for the style
            tcBorderType.SetAttributeValue(DocxNamespace.Main + "val", border.Style.GetEnumName());

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
                _ => 2
            };

            // The sz attribute is used for the border size
            tcBorderType.SetAttributeValue(DocxNamespace.Main + "sz", size);

            // The space attribute is used for the cell spacing (probably '0')
            tcBorderType.SetAttributeValue(DocxNamespace.Main + "space", border.SpacingOffset);

            // The color attribute is used for the border color
            tcBorderType.SetAttributeValue(DocxNamespace.Main + "color", border.Color.ToHex());
        }

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected override void OnDocumentOwnerChanged(DocX previousValue, DocX newValue)
        {
            base.OnDocumentOwnerChanged(previousValue, newValue);
            PackagePart = Row?.PackagePart;
        }

    }
}