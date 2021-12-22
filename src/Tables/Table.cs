using DXPlus.Helpers;
using DXPlus.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a Table in a document {tbl}
    /// </summary>
    public class Table : Block, IEquatable<Table>
    {
        private string customTableDesignName;
        private TableDesign tableDesign;

        /// <summary>
        /// Public constructor to create an empty table
        /// </summary>
        public Table() : this(null, null, CreateEmptyTableXml())
        {
        }

        /// <summary>
        /// Public constructor to create an empty uniform table
        /// </summary>
        /// <param name="rows">Rows</param>
        /// <param name="columns">Columns</param>
        public Table(int rows, int columns) : this()
        {
            if (rows <= 0)
                throw new ArgumentOutOfRangeException("Rows must be >= 1", nameof(rows));
            if (columns <= 0)
                throw new ArgumentOutOfRangeException("Columns must be >= 1", nameof(columns));

            WriteTableConditionalFormat(TblPr, TableConditionalFormatting.None);

            var grid = Xml.GetOrAddElement(Namespace.Main + "tblGrid");
            double width = double.NaN;
            for (int i = 0; i < columns; i++)
            {
                grid.Add(new XElement(Namespace.Main + "gridCol", new XAttribute(Namespace.Main + "w", width)));
            }

            for (int i = 0; i < rows; i++)
            {
                var row = new XElement(Namespace.Main + "tr");
                foreach (var cw in DefaultColumnWidths)
                    row.Add(CreateTableCell(cw));
                Xml.Add(row);
            }
        }

        /// <summary>
        /// Internal constructor for the table - wraps a {tbl} element.
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="packagePart">Package owner</param>
        /// <param name="xml">XML fragment representing the table</param>
        internal Table(IDocument document, PackagePart packagePart, XElement xml) : base(document, packagePart, xml)
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

            set => WriteTableConditionalFormat(TblPr, value);
        }

        /// <summary>
        /// How the preferred table width is expressed.
        /// </summary>
        public TableWidthUnit PreferredTableWidthUnit => Enum.TryParse<TableWidthUnit>(TblPr.Element(Namespace.Main + "tblW")?
                        .AttributeValue(Namespace.Main + "type"), ignoreCase: true, out var tbw) ? tbw : TableWidthUnit.Auto;

        /// <summary>
        /// Preferred table width.
        /// </summary>
        public double PreferredTableWidth => double.TryParse(TblPr.Element(Namespace.Main + "tblW")?.AttributeValue(Namespace.Main + "w"), out var d) ? d : 0;

        /// <summary>
        /// Sets the table width
        /// </summary>
        /// <param name="unitType">Units</param>
        /// <param name="value">Expressed width in specified units</param>
        public void SetTableWidth(TableWidthUnit unitType, double? value)
        {
            XElement tblW = TblPr.GetOrAddElement(Namespace.Main + "tblW");
            if (unitType == TableWidthUnit.None || value == null || value < 0)
            {
                tblW.Remove();
                return;
            }

            tblW.SetAttributeValue(Namespace.Main + "type", unitType.GetEnumName());
            if (unitType == TableWidthUnit.Auto)
                value = 0;

            if (unitType == TableWidthUnit.Percentage)
                tblW.SetAttributeValue(Namespace.Main + "w", value.Value + "%");
            else
                tblW.SetAttributeValue(Namespace.Main + "w", value.Value);
        }

        /// <summary>
        /// True if the table will auto-fit the contents. This corresponds to the {tblLayout} value of the table properties.
        /// </summary>
        public bool AutoFit
        {
            get => string.Compare(TblPr.Element(Namespace.Main + "tblLayout")?.AttributeValue(Namespace.Main + "type"), "autofit", StringComparison.CurrentCultureIgnoreCase) == 0;
            set => TblPr.GetOrAddElement(Namespace.Main + "tblLayout").SetAttributeValue(Namespace.Main + "type", value ? "autofit" : "fixed");
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
        /// Indentation in dxa units
        /// </summary>
        public double? Indent
        {
            get
            {
                XAttribute value = TblPr.Element(Name.TableIndent)?
                                        .Attribute(Namespace.Main + "w");
                if (value != null && double.TryParse(value.Value, out var indentUnits))
                    return indentUnits;

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
                    tblIndent.SetAttributeValue(Namespace.Main + "type", "dxa");
                    tblIndent.SetAttributeValue(Namespace.Main + "w", value);
                }
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
                .Select(c => double.TryParse(c.AttributeValue(Namespace.Main + "w"), out var dbl) ? dbl : double.NaN);

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
        protected override void OnDocumentOwnerChanged()
        {
            if (ColumnCount == 0 || !Rows.Any())
                throw new Exception("Cannot add empty table to document.");

            // Fixup any unsized columns
            if (DefaultColumnWidths.Any(dc => double.IsNaN(dc)))
            {
                OnSetColumnWidths(DefaultColumnWidths.ToArray());
            }

            // Add any required styles
            ApplyTableStyleToDocumentOwner();

            // Set the document/package for each row.
            Rows.ToList().ForEach(r => r.SetOwner(Document, PackagePart));
        }

        /// <summary>
        /// This ensures the owning document has the table style applied.
        /// </summary>
        private void ApplyTableStyleToDocumentOwner()
        {
            if (Document == null) return;

            string designName = TblPr.Element(Namespace.Main + "tblStyle").GetVal();
            if (string.IsNullOrWhiteSpace(designName) || string.Compare(designName, "none", StringComparison.InvariantCultureIgnoreCase)==0)
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
        public IEnumerable<TableRow> Rows => Xml.Elements(Namespace.Main + "tr").Select(r => new TableRow(this, r));

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
            if (index < 0 || index > ColumnCount)
                throw new ArgumentOutOfRangeException("Insert position must fall within existing column range", nameof(index));

            bool insertAtEnd = (index == ColumnCount);

            // Create a new column by splitting the last column in half.
            var grid = Xml.GetOrAddElement(Namespace.Main + "tblGrid");
            var gridColumns = grid.Elements(Namespace.Main + "gridCol").ToList();

            XElement existingGridCol = insertAtEnd ? gridColumns.Last() : gridColumns[index];
            double width = double.Parse(existingGridCol.AttributeValue(Namespace.Main + "w"));
            existingGridCol.SetAttributeValue(Namespace.Main + "w", width);

            // Add a new <gridCol> definition
            var newColumnDef = new XElement(Namespace.Main + "gridCol", new XAttribute(Namespace.Main + "w", width));
            if (insertAtEnd)
                grid.Add(newColumnDef);
            else
                existingGridCol.AddBeforeSelf(newColumnDef);

            // Now add a blank column to each row and update the cell widths.
            var rows = Rows.ToList();
            foreach (var row in rows)
            {
                var cell = CreateTableCell(width);
                var cells = row.Cells;
                if (insertAtEnd)
                {
                    cells[index - 1].Xml.AddAfterSelf(cell);
                }
                else
                {
                    cells[index].Xml.AddBeforeSelf(cell);
                }
            }
        }

        /// <summary>
        /// Insert a blank row at the end of this table.
        /// </summary>
        public TableRow AddRow()
        {
            return InsertRow(Rows.Count());
        }

        /// <summary>
        /// Insert a row into this table.
        /// </summary>
        public TableRow InsertRow(int index)
        {
            var rows = Rows.ToList();

            if (index < 0 || index > rows.Count) throw new ArgumentOutOfRangeException(nameof(index));

            var content = new List<XElement>();
            var columnWidths = DefaultColumnWidths.ToList();
            for (int i = 0; i < ColumnCount; i++)
            {
                double? width = null;
                if (columnWidths.Count > i)
                {
                    width = columnWidths[i];
                }

                content.Add(CreateTableCell(width));
            }

            return InsertRow(index, content);
        }

        /// <summary>
        /// Inserts a row using the passed elements at the given index.
        /// </summary>
        /// <param name="index"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        private TableRow InsertRow(int index, IEnumerable<XElement> content)
        {
            if (content == null)
                throw new ArgumentNullException(nameof(content));
            if (!content.Any())
                throw new ArgumentException("Must have content to insert a row.", nameof(content));

            var rows = Rows.ToList();
            if (index < 0 || index > rows.Count)
                throw new ArgumentOutOfRangeException(nameof(index));

            var row = new TableRow(this, new XElement(Namespace.Main + "tr", content));

            if (index == rows.Count)
                rows[^1].Xml.AddAfterSelf(row.Xml);
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

            if (endRow > Rows.Count())
                throw new IndexOutOfRangeException(nameof(count));

            var startRowElement = Rows.ElementAt(startRow).Cells[columnIndex].Xml;

            // Move the content over and add vMerge to each row cell
            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                TableCell cell = Rows.ElementAt(rowIndex).Cells[columnIndex];
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

            // Remove the column from default columns.
            Xml.Element(Namespace.Main + "tblGrid").Elements(Namespace.Main + "gridCol").ElementAt(index).Remove();

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
            List<TableRow> rows = Rows.ToList();
            if (index < 0 || index > rows.Count - 1)
            {
                throw new IndexOutOfRangeException();
            }

            rows[index].Xml.Remove();
            if (!Rows.Any()) // use real property
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
        public void SetColumnWidths(IEnumerable<double> widths)
        {
            if (widths.Count() != ColumnCount)
            {
                throw new ArgumentOutOfRangeException(nameof(widths), "Must supply widths for each column.");
            }

            OnSetColumnWidths(widths.ToArray());
        }

        /// <summary>
        /// Sets the column width for the given index.
        /// </summary>
        /// <param name="index">Column index</param>
        /// <param name="width">Column width</param>
        public void SetColumnWidth(int index, double width)
        {
            double[] columnWidths = DefaultColumnWidths?.ToArray();
            if (columnWidths == null || index > columnWidths.Length - 1)
            {
                if (!Rows.Any())
                {
                    throw new InvalidOperationException("Must have at least one row to determine column widths.");
                }

                columnWidths = Rows.ToList()[^1].Cells.Select(c => c.Width ?? 0).ToArray();
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
        private void OnSetColumnWidths(double[] columnWidths)
        {
            // Fill in any missing values.
            if (columnWidths.Contains(double.NaN))
            {
                double totalSpace;

                if (PreferredTableWidthUnit == TableWidthUnit.Dxa)
                    totalSpace = PreferredTableWidth;
                else
                {
                    totalSpace = Document.Sections.First().Properties.AdjustedPageWidth;
                    if (PreferredTableWidthUnit == TableWidthUnit.Percentage)
                    {
                        totalSpace *= PreferredTableWidth;
                    }
                }

                double usedSpace = columnWidths.Where(c => !double.IsNaN(c)).Sum();
                double eachColumn = (totalSpace - usedSpace) / columnWidths.Count(double.IsNaN);
                for (int i = 0; i < columnWidths.Length; i++)
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
                    new XAttribute(Namespace.Main + "w", width)));
            }

            // Reset cell widths
            foreach (TableRow row in Rows)
            {
                row.SetColumnWidths(columnWidths);
            }

            // Set to fixed sizing
            AutoFit = false;
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
                            .AttributeValue(Namespace.Main + "w"), out double result) ? result : null;
        }

        /// <summary>
        /// Set the specified cell margin for the table-level.
        /// </summary>
        /// <param name="type">The side of the cell margin.</param>
        /// <param name="margin">The value for the specified cell margin in dxa units.</param>
        public void SetDefaultCellMargin(TableCellMarginType type, double? margin)
        {
            if (margin != null)
            {
                XElement cellMargin = TblPr.GetOrAddElement(Namespace.Main + "tblCellMar")
                    .GetOrAddElement(Namespace.Main + type.GetEnumName());
                cellMargin.SetAttributeValue(Namespace.Main + "w", margin);
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
        private XElement GetOrCreateTablePropertiesSection() => Xml.GetOrAddElement(Namespace.Main + "tblPr");

        /// <summary>
        /// Write the element children for the TableConditionalFormatting
        /// </summary>
        /// <param name="tblPr"></param>
        /// <param name="format"></param>
        private static void WriteTableConditionalFormat(XElement tblPr, TableConditionalFormatting format)
        {
            var e = tblPr.GetOrAddElement(Namespace.Main + "tblLook");
            e.RemoveAttributes();

            e.Add(
                new XAttribute(Namespace.Main + "firstColumn", format.HasFlag(TableConditionalFormatting.FirstColumn) ? 1 : 0),
                new XAttribute(Namespace.Main + "lastColumn", format.HasFlag(TableConditionalFormatting.LastColumn) ? 1 : 0),
                new XAttribute(Namespace.Main + "firstRow", format.HasFlag(TableConditionalFormatting.FirstRow) ? 1 : 0),
                new XAttribute(Namespace.Main + "lastRow", format.HasFlag(TableConditionalFormatting.LastRow) ? 1 : 0),
                new XAttribute(Namespace.Main + "noHBand", format.HasFlag(TableConditionalFormatting.NoRowBand) ? 1 : 0),
                new XAttribute(Namespace.Main + "noVBand", format.HasFlag(TableConditionalFormatting.NoColumnBand) ? 1 : 0),
                new XAttribute(Name.MainVal, format.ToHex(4)));
        }

        /// <summary>
        /// Create and return the default {tbl} root element
        /// </summary>
        private static XElement CreateEmptyTableXml() => new XElement(Name.Table,
                new XElement(Namespace.Main + "tblPr",
                    new XElement(Namespace.Main + "tblStyle",
                        new XAttribute(Name.MainVal, "None")),
                    new XElement(Namespace.Main + "tblW",
                        new XAttribute(Namespace.Main + "type", "auto"),
                        new XAttribute(Namespace.Main + "w", 0))),
                new XElement(Namespace.Main + "tblGrid"));

        /// <summary>
        /// Create and return a cell of a table
        /// </summary>
        /// <param name="width">Optional width of the cell in dxa units</param>
        private static XElement CreateTableCell(double? width = null)
        {
            string type = width == null ? "auto" : "dxa";
            width ??= 0;

            return new XElement(Namespace.Main + "tc",
                    new XElement(Namespace.Main + "tcPr",
                        new XElement(Namespace.Main + "tcW",
                            new XAttribute(Namespace.Main + "type", type),
                            new XAttribute(Namespace.Main + "w", width.Value)
                        )),
                    // always has an empty paragraph
                    new XElement(Name.Paragraph,
                        new XAttribute(Name.ParagraphId,
                            HelperFunctions.GenerateHexId()))
            );
        }

        /// <summary>
        /// Determines equality for a table object
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(Table other)
        {
            if (other == null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Xml == other.Xml;
        }
    }
}