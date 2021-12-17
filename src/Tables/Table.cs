using DXPlus.Helpers;
using DXPlus.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a Table in a document.
    /// </summary>
    public class Table : Block
    {
        private string customTableDesignName;
        private TableDesign tableDesign;

        /// <summary>
        /// Public constructor to create an empty table
        /// </summary>
        /// <param name="rows">Rows</param>
        /// <param name="columns">Columns</param>
        public Table(int rows, int columns) : this(null, TableHelpers.CreateTable(rows, columns))
        {
        }

        /// <summary>
        /// Create a table from a text array of rows, columns
        /// </summary>
        /// <param name="values"></param>
        public Table(string[,] values) : this(null, TableHelpers.CreateTable(values.GetLength(0), values.GetLength(1)))
        {
            int rowIndex = 0, colIndex = 0;
            foreach (var row in Rows)
            {
                foreach (var cell in row.Cells)
                {
                    cell.Text = values[rowIndex, colIndex];
                    colIndex++;
                }
                rowIndex++;
            }
        }

        /// <summary>
        /// Constructor for the table
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">XML fragment representing the table</param>
        internal Table(IDocument document, XElement xml) : base(document, xml)
        {
            if (xml == null)
                throw new ArgumentNullException(nameof(xml));
            if (xml.Name != Name.Table)
                throw new ArgumentException($"Root element must be {Name.Table}", nameof(xml));

            var style = TblPr?.Element(Namespace.Main + "tblStyle");
            if (style == null)
            {
                tableDesign = TableDesign.None;
            }
            else
            {
                tableDesign = style.TryGetEnumValue(out TableDesign tdr) ? tdr : TableDesign.Custom;
                if (tableDesign == TableDesign.Custom)
                {
                    customTableDesignName = style.GetVal();
                }
            }
        }

        /// <summary>
        /// Get the table properties
        /// </summary>
        private XElement TblPr => GetOrCreateTablePropertiesSection();

        /// <summary>
        /// The conditional formatting applied to the table
        /// </summary>
        public TableConditionalFormatting ConditionalFormatting
        {
            get => Enum.TryParse<TableConditionalFormatting>(TblPr.Element(Namespace.Main + "tblLook")?.GetVal(),
                out var tcf)
                ? tcf
                : TableConditionalFormatting.None;

            set => TableHelpers.WriteTableConditionalFormat(TblPr, value);
        }

        /// <summary>
        /// Specifies the alignment of the current table with respect to the text margins in the current section
        /// </summary>
        public Alignment Alignment
        {
            get => TblPr.Element(Name.ParagraphAlignment)
                    .GetVal().TryGetEnumValue(out Alignment result)
                    ? result
                    : Alignment.Left;

            set => TblPr.GetOrAddElement(Name.ParagraphAlignment)
                    .SetAttributeValue(Name.MainVal,
                        value.GetEnumName());
        }

        /// <summary>
        /// Indentation in pixels
        /// </summary>
        public double? Indent
        {
            get
            {
                XAttribute value = TblPr.Element(Name.TableIndent)?
                                        .Attribute(Namespace.Main + "w");
                if (value != null && double.TryParse(value.Value, out var indentUnits))
                    return indentUnits / TableHelpers.UnitConversion;

                value?.Remove();
                return null;
            }
            set
            {
                XElement tblIndent = TblPr.GetOrAddElement(Name.TableIndent);
                if (value is null or < 0)
                {
                    tblIndent.Remove();
                }
                else
                {
                    tblIndent.SetAttributeValue(Namespace.Main + "type", "dxa"); // Widths in 20th/pt.
                    tblIndent.SetAttributeValue(Namespace.Main + "w", value * TableHelpers.UnitConversion);
                }
            }
        }

        /// <summary>
        /// Auto size this table according to some rule.
        /// </summary>
        public AutoFit AutoFit
        {
            get
            {
                string preferredTableWidth = TblPr.Element(Namespace.Main + "tblW").AttributeValue(Namespace.Main + "type", "auto");

                if (string.Equals("auto", preferredTableWidth, StringComparison.CurrentCultureIgnoreCase))
                {
                    return Xml.LocalNameDescendants("tcW")
                              .All(e => e.AttributeValue(Namespace.Main + "type", "auto") == "auto")
                        ? AutoFit.Contents
                        : AutoFit.ColumnWidth;
                }

                if (string.Equals("pct", preferredTableWidth, StringComparison.CurrentCultureIgnoreCase))
                {
                    return AutoFit.Window;
                }

                if (string.Equals("dxa", preferredTableWidth, StringComparison.CurrentCultureIgnoreCase))
                {
                    return AutoFit.Fixed;
                }

                throw new InvalidOperationException($"Unsupported table width preference - {preferredTableWidth}.");
            }

            set
            {
                string preferredTableWidth = null, preferredColWidth = null;

                switch (value)
                {
                    case AutoFit.ColumnWidth:
                        preferredTableWidth = "auto";
                        preferredColWidth = "dxa";
                        // Disable "Automatically resize to fit contents" option
                        TblPr.SetAttributeValue(Namespace.Main + "tblLayout", Namespace.Main + "type", "fixed");
                        break;

                    case AutoFit.Contents:
                        preferredTableWidth = "auto";
                        preferredColWidth = "auto";
                        break;

                    case AutoFit.Window:
                        preferredTableWidth = "pct";
                        preferredColWidth = "pct";
                        break;

                    case AutoFit.Fixed:
                        preferredTableWidth = "dxa";
                        preferredColWidth = "dxa";

                        XElement tblLayout = TblPr?.Element(Namespace.Main + "tblLayout");
                        if (tblLayout == null)
                        {
                            XElement tmp = TblPr?.Element(Namespace.Main + "tblInd") ??
                                      TblPr?.Element(Namespace.Main + "tblW");
                            if (tmp != null)
                            {
                                tmp.AddAfterSelf(new XElement(Namespace.Main + "tblLayout"));
                                TblPr.Element(Namespace.Main + "tblLayout")
                                    ?.SetAttributeValue(Namespace.Main + "type", "fixed");
                                TblPr.Element(Namespace.Main + "tblW")
                                    ?.SetAttributeValue(Namespace.Main + "w", DefaultColumnWidths.Sum());
                            }
                            break;
                        }

                        foreach (XAttribute type in Xml.LocalNameDescendants("tblLayout")
                            .Attributes(Namespace.Main + "type"))
                        {
                            type.Value = "fixed";
                        }

                        // Set the table width based on the column widths
                        TblPr.GetOrAddElement(Namespace.Main + "tblW")?
                             .SetAttributeValue(Namespace.Main + "w", DefaultColumnWidths.Sum());
                        break;
                }

                // Set preferred width exception type
                foreach (var type in Xml.LocalNameDescendants("tblW"))
                    type.SetAttributeValue(Namespace.Main + "type", preferredTableWidth);

                // Set preferred table cell width type
                foreach (var type in Xml.LocalNameDescendants("tcW"))
                    type.SetAttributeValue(Namespace.Main + "type", preferredColWidth);
            }
        }

        /// <summary>
        /// Returns the number of columns in this table.
        /// </summary>
        public int ColumnCount => DefaultColumnWidths.Count();

        /// <summary>
        /// Gets the default column widths for this table. An individual row can override this
        /// by providing an explicit width on the cell.
        /// </summary>
        public IEnumerable<double> DefaultColumnWidths =>
            Xml.Element(Namespace.Main + "tblGrid")?
                .Elements(Namespace.Main + "gridCol")
                .Select(c => Convert.ToDouble(c.AttributeValue(Namespace.Main + "w")));

        /// <summary>
        /// Custom Table Style name
        /// </summary>
        public string CustomTableDesignName
        {
            get => customTableDesignName;
            set
            {
                customTableDesignName = value;
                tableDesign = TableDesign.Custom;
                TblPr.GetOrAddElement(Namespace.Main + "tblStyle")
                     .SetAttributeValue(Name.MainVal, value);
            }
        }

        /// <summary>
        /// The design\style to apply to this table.
        /// </summary>
        public TableDesign Design
        {
            get => tableDesign;
            set
            {
                if (value == TableDesign.Custom)
                {
                    throw new ArgumentOutOfRangeException(nameof(Design), $"Cannot set custom design value - use {CustomTableDesignName} property instead.");
                }

                tableDesign = value;

                XElement style = TblPr.GetOrAddElement(Namespace.Main + "tblStyle");
                if (tableDesign == TableDesign.None)
                {
                    style?.Remove();
                }
                else
                {
                    style.SetAttributeValue(Name.MainVal, tableDesign.GetEnumName());
                }

                ApplyTableStyleToDocumentOwner();
            }
        }

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected override void OnDocumentOwnerChanged(IDocument previousValue, IDocument newValue)
        {
            base.OnDocumentOwnerChanged(previousValue, newValue);

            if (newValue is Document doc && Xml != null)
            {
                ApplyTableStyleToDocumentOwner();
                Rows.ToList().ForEach(r => r.Document = doc);
            }
        }

        /// <summary>
        /// Called when the package part is changed.
        /// </summary>
        protected override void OnPackagePartChanged(PackagePart previousValue, PackagePart newValue)
        {
            base.OnPackagePartChanged(previousValue, newValue);
            Rows.ToList().ForEach(r => r.PackagePart = newValue);
        }

        /// <summary>
        /// This ensures the owning document has the table style applied.
        /// </summary>
        private void ApplyTableStyleToDocumentOwner()
        {
            if (Document == null)
                return;

            string designName = TblPr.Element(Namespace.Main + "tblStyle").GetVal();
            if (string.IsNullOrWhiteSpace(designName))
                return;

            if (!Document.Styles.HasStyle(designName, StyleType.Table))
            {
                var styleElement = Resource.DefaultTableStyles()
                                        .Descendants()
                                        .FindByAttrVal(Namespace.Main + "styleId", designName);
                Document.Styles.Add(styleElement);
            }
        }

        /// <summary>
        /// Get all of the Hyperlinks in this Table.
        /// </summary>
        public IEnumerable<Hyperlink> Hyperlinks => Rows.SelectMany(r => r.Cells).SelectMany(c => c.Hyperlinks);

        /// <summary>
        /// Returns Paragraphs inside this container.
        /// </summary>
        public IEnumerable<Paragraph> Paragraphs => Rows.SelectMany(r => r.Cells).SelectMany(c => c.Paragraphs);

        /// <summary>
        /// Returns Pictures in this container.
        /// </summary>
        public IEnumerable<Picture> Pictures => Rows.SelectMany(r => r.Cells).SelectMany(c => c.Pictures);

        /// <summary>
        /// Returns a list of rows in this table.
        /// </summary>
        public IEnumerable<Row> Rows => Xml.Elements(Namespace.Main + "tr").Select(r => new Row(this, r));

        /// <summary>
        /// Gets or Sets the value of the Table Caption (Alternate Text Title) of this table.
        /// </summary>
        public string TableCaption
        {
            get => TblPr.Element(Namespace.Main + "tblCaption")?.GetVal() ?? string.Empty;
            set
            {
                TblPr?.Descendants(Namespace.Main + "tblCaption").FirstOrDefault()?.Remove();
                TblPr?.Add(new XElement(Namespace.Main + "tblCaption", new XAttribute(Name.MainVal, value)));
            }
        }

        /// <summary>
        /// Gets or Sets the value of the Table Description (Alternate Text Description) of this table.
        /// </summary>
        public string TableDescription
        {
            get => TblPr.Element(Namespace.Main + "tblDescription")?.GetVal() ?? string.Empty;
            set
            {
                TblPr.Descendants(Namespace.Main + "tblDescription").FirstOrDefault()?.Remove();
                TblPr.Add(new XElement(Namespace.Main + "tblDescription", new XAttribute(Name.MainVal, value)));
            }
        }

        /// <summary>
        /// Add a column on the end of the table. Each row is given an empty cell.
        /// </summary>
        public void AddColumn()
        {
            InsertColumn(ColumnCount);
        }

        /// <summary>
        /// Insert a column into a table. All rows are given an empty cell.
        /// </summary>
        /// <param name="index">The index to insert the column at.</param>
        public void InsertColumn(int index)
        {
            // Create a new column by reducing the width of the other columns
            var grid = Xml.GetOrAddElement(Namespace.Main + "tblGrid");
            var columns = grid.Elements(Namespace.Main + "gridCol").ToList();
            int currentSize = columns.Sum(c => int.TryParse(c.AttributeValue(Namespace.Main + "w"), out int val) ? val : 0);
            if (currentSize <= 0 || currentSize > TableHelpers.MaxTableWidth)
                currentSize = TableHelpers.MaxTableWidth;

            var rows = Rows.ToList();
            int columnCount = Math.Max(columns.Count, rows.Max(r => r.ColumnCount));
            var columnWidths = TableHelpers.CalculateProportionalWidths(columnCount, currentSize);

            columns.ForEach(c => c.Remove());
            foreach (int width in columnWidths)
            {
                grid.Add(new XElement(Namespace.Main + "gridCol", new XAttribute(Namespace.Main + "w", width * TableHelpers.UnitConversion)));
            }

            // Now add a blank column to each row and update the cell widths.
            if (rows.Count>0)
            {
                // Add a new Cell to each row
                foreach (var row in rows)
                {
                    var cell = TableHelpers.CreateTableCell();
                    var cells = row.Cells;
                    if (cells.Count == index)
                        cells[index - 1].Xml.AddAfterSelf(cell);
                    else
                        cells[index].Xml.AddBeforeSelf(cell);
                    for (var i = 0; i < cells.Count; i++)
                        cells[i].Width = columnWidths[i];
                }
            }
        }

        /// <summary>
        /// Insert a blank row at the end of this table.
        /// </summary>
        public Row AddRow()
        {
            return InsertRow(Rows.Count());
        }

        /// <summary>
        /// Insert a row into this table.
        /// </summary>
        public Row InsertRow(int index)
        {
            var rows = Rows.ToList();

            if (index <= 0 || index > rows.Count) throw new ArgumentOutOfRangeException(nameof(index));

            var content = new List<XElement>();
            var columnWidths = DefaultColumnWidths.ToList();
            for (int i = 0; i < ColumnCount; i++)
            {
                double? width = null;
                if (columnWidths.Count > i)
                {
                    width = columnWidths[i];
                }

                content.Add(TableHelpers.CreateTableCell(width));
            }

            return InsertRow(index, content);
        }

        /// <summary>
        /// Inserts a row using the passed elements at the given index.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        private Row InsertRow(int index, IEnumerable<XElement> content)
        {
            if (content == null)
                throw new ArgumentNullException(nameof(content));
            if (!content.Any())
                throw new ArgumentException("Must have content to insert a row.", nameof(content));

            var rows = Rows;
            if (index < 0 || index > rows.Count)
                throw new ArgumentOutOfRangeException(nameof(index));

            var row = new Row(this, new XElement(Namespace.Main + "tr", content));

            if (index == rows.Count)
                rows.Last().Xml.AddAfterSelf(row.Xml);
            else
                rows[index].Xml.AddBeforeSelf(row.Xml);

            return row;
        }

        /// <summary>
        /// Merge cells in given column starting with startRow.
        /// </summary>
        public void MergeCellsInColumn(int columnIndex, int startRow, int count)
        {
            int endRow = startRow + count - 1;
            if (columnIndex < 0 || columnIndex >= ColumnCount)
                throw new IndexOutOfRangeException(nameof(columnIndex));

            if (startRow < 0 || startRow >= endRow)
                throw new IndexOutOfRangeException(nameof(startRow));

            if (endRow > Rows.Count)
                throw new IndexOutOfRangeException(nameof(count));

            var startRowElement = Rows[startRow].Cells[columnIndex].Xml;

            // Move the content over and add vMerge to each row cell
            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                Cell cell = Rows[rowIndex].Cells[columnIndex];
                XElement vMerge = cell.Xml.GetOrAddElement(Namespace.Main + "tcPr")
                            .GetOrAddElement(Namespace.Main + "vMerge");

                if (rowIndex == startRow)
                {
                    vMerge.SetAttributeValue(Name.MainVal, "restart");
                }
                else
                {
                    List<XElement> paragraphs = cell.Xml.Elements(Name.Paragraph).ToList();
                    if (paragraphs.Count > 0)
                    {
                        startRowElement.Add(paragraphs);
                        cell.Text = null;
                    }
                }
            }
        }

        /// <summary>
        /// Remove this Table from any document owner.
        /// We leave the Document in place.
        /// </summary>
        public void Remove()
        {
            if (Xml.Parent != null)
            {
                Xml.Remove();
            }
        }

        /// <summary>
        /// Remove a column from this Table.
        /// </summary>
        /// <param name="index">The column to remove.</param>
        public void RemoveColumn(int index)
        {
            if (index < 0 || index > ColumnCount - 1)
                throw new IndexOutOfRangeException();

            foreach (var row in Rows)
                row.Cells[index].Xml.Remove();

            if (ColumnCount == 0)
                Remove();
        }

        /// <summary>
        /// Remove a row from this Table.
        /// </summary>
        /// <param name="index">The row to remove.</param>
        public void RemoveRow(int index)
        {
            List<Row> rows = Rows;
            if (index < 0 || index > rows.Count - 1)
            {
                throw new IndexOutOfRangeException();
            }

            rows[index].Xml.Remove();
            if (Rows.Count == 0) // use real property
                Remove();
        }

        /// <summary>
        /// Set outside borders to the given style
        /// </summary>
        public void SetOutsideBorders(BorderStyle style, Color color, double? spacing = 1, double size = 2)
        {
            SetBorder(TableBorderType.Top, style, color, spacing, size);
            SetBorder(TableBorderType.Left, style, color, spacing, size);
            SetBorder(TableBorderType.Right, style, color, spacing, size);
            SetBorder(TableBorderType.Bottom, style, color, spacing, size);
        }

        /// <summary>
        /// Set all inside borders to the given style
        /// </summary>
        public void SetInsideBorders(BorderStyle style, Color color, double? spacing = 1, double size = 2)
        {
            SetBorder(TableBorderType.InsideV, style, color, spacing, size);
            SetBorder(TableBorderType.InsideH, style, color, spacing, size);
        }

        /// <summary>
        /// Set a table border
        /// </summary>
        /// <exception cref="InvalidEnumArgumentException"></exception>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public void SetBorder(TableBorderType borderType, BorderStyle style, Color color, double? spacing = 1, double size = 2)
        {
            if (size is < 2 or > 96)
                throw new ArgumentOutOfRangeException(nameof(Size));
            if (!Enum.IsDefined(typeof(ParagraphBorderType), borderType))
                throw new InvalidEnumArgumentException(nameof(borderType), (int)borderType, typeof(ParagraphBorderType));

            // Set the border style
            TblPr.Element(Namespace.Main + "tblBorders")?
                 .Element(Namespace.Main + borderType.GetEnumName())?.Remove();

            if (style == BorderStyle.None)
                return;

            var tblBorders = TblPr.GetOrAddElement(Namespace.Main + "tblBorders");
            var borderXml = new XElement(Namespace.Main + borderType.GetEnumName(),
                new XAttribute(Name.MainVal, style.GetEnumName()),
                new XAttribute(Name.Size, size));
            if (color != Color.Empty)
                borderXml.Add(new XAttribute(Name.Color, color.ToHex()));
            if (spacing != null)
                borderXml.Add(new XAttribute(Namespace.Main + "space", spacing));

            tblBorders.Add(borderXml);
        }

        /// <summary>
        /// Supply all the column widths
        /// </summary>
        /// <param name="widths"></param>
        public void SetColumnWidths(IList<double> widths)
        {
            if (widths.Count != ColumnCount)
            {
                throw new ArgumentOutOfRangeException(nameof(widths), "Must supply widths for each column.");
            }

            OnSetColumnWidths(widths);
        }

        /// <summary>
        /// Sets the column width for the given index.
        /// </summary>
        /// <param name="index">Column index</param>
        /// <param name="width">Column width</param>
        public void SetColumnWidth(int index, double width)
        {
            List<double> columnWidths = DefaultColumnWidths?.ToList();
            if (columnWidths == null || index > columnWidths.Count - 1)
            {
                if (Rows.Count == 0)
                {
                    throw new InvalidOperationException("Must have at least one row to determine column widths.");
                }

                columnWidths = Rows[^1].Cells.Select(c => c.Width ?? 0).ToList();
            }

            if (width >= 0)
            {
                columnWidths[index] = width;
            }

            OnSetColumnWidths(columnWidths);
        }

        /// <summary>
        /// Rewrite the tbl/tblGrid section with new column widths
        /// </summary>
        /// <param name="columnWidths">Set of widths</param>
        private void OnSetColumnWidths(IList<double> columnWidths)
        {
            // Fill in any missing values.
            if (columnWidths.Contains(double.NaN))
            {
                SectionProperties section = Document.Sections.First().Properties;
                double pageWidth = section.PageWidth - section.LeftMargin - section.RightMargin;
                double usedSpace = columnWidths.Where(c => !double.IsNaN(c)).Sum();
                double eachColumn = (pageWidth - usedSpace) / columnWidths.Count(double.IsNaN);
                for (int i = 0; i < columnWidths.Count; i++)
                {
                    if (double.IsNaN(columnWidths[i]))
                    {
                        columnWidths[i] = eachColumn;
                    }
                }
            }

            // Replace the columns with the new values.
            XElement grid = Xml.Element(Namespace.Main + "tblGrid");
            if (grid != null)
            {
                grid.RemoveAll();
            }
            else
            {
                grid = new XElement(Namespace.Main + "tblGrid");
                TblPr.AddAfterSelf(grid);
            }

            foreach (double width in columnWidths)
            {
                grid.Add(new XElement(Namespace.Main + "gridCol",
                    new XAttribute(Namespace.Main + "w", width * TableHelpers.UnitConversion)));
            }

            // Reset cell widths
            foreach (Row row in Rows)
            {
                row.SetColumnWidths(columnWidths);
            }

            // Set to fixed sizing
            AutoFit = AutoFit.Fixed;
        }

        /// <summary>
        /// Gets the table margin value in pixels for the specified margin
        /// </summary>
        /// <param name="type">Table margin type</param>
        /// <returns>The value for the specified margin in pixels, null if it's not set.</returns>
        public double? GetDefaultCellMargin(TableCellMarginType type)
        {
            return double.TryParse(TblPr.Element(Namespace.Main + "tblCellMar")?
                            .Element(Namespace.Main + type.GetEnumName())?
                            .AttributeValue(Namespace.Main + "w"), out double result)
                            ? (double?)result / TableHelpers.UnitConversion
                            : null;
        }

        /// <summary>
        /// Set the specified cell margin for the table-level.
        /// </summary>
        /// <param name="type">The side of the cell margin.</param>
        /// <param name="margin">The value for the specified cell margin.</param>
        public void SetDefaultCellMargin(TableCellMarginType type, double? margin)
        {
            if (margin != null)
            {
                XElement cellMargin = TblPr.GetOrAddElement(Namespace.Main + "tblCellMar")
                    .GetOrAddElement(Namespace.Main + type.GetEnumName());
                cellMargin.SetAttributeValue(Namespace.Main + "w", margin * TableHelpers.UnitConversion);
                cellMargin.SetAttributeValue(Namespace.Main + "type", "dxa");
            }
            else
            {
                XElement margins = TblPr.Element(Namespace.Main + "tblCellMar");
                margins?.Element(Namespace.Main + type.GetEnumName())?.Remove();
                if (margins?.IsEmpty == true)
                {
                    margins.Remove();
                }
            }
        }

        /// <summary>
        /// Retrieves or create the table properties (tblPr) section.
        /// </summary>
        /// <returns>The w:tbl/tblPr element.</returns>
        private XElement GetOrCreateTablePropertiesSection()
        {
            return Xml.GetOrAddElement(Namespace.Main + "tblPr");
        }
    }
}