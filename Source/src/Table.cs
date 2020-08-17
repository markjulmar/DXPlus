using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{

    /// <summary>
    /// Represents a Table in a document.
    /// </summary>
    public class Table : InsertBeforeOrAfter
    {
        private string customTableDesignName;
        private TableDesign tableDesign;
        private Alignment alignment;
        private AutoFit autoFit;
        private int? cachedColCount;
        private double[] columnWidths;

        /// <summary>
        /// Constructor for the table
        /// </summary>
        /// <param name="document">Document owner</param>
        /// <param name="xml">XML fragment representing the table</param>
        internal Table(DocX document, XElement xml) : base(document, xml)
        {
            autoFit = AutoFit.ColumnWidth;
            packagePart = document.packagePart;

            var style = TblPr?.Element(DocxNamespace.Main + "tblStyle");
            if (style != null)
            {
                var val = style.GetValAttr();
                tableDesign = val != null
                    ? Enum.TryParse<TableDesign>(val.Value.Replace("-", ""), out TableDesign result) ? result : TableDesign.Custom
                    : TableDesign.None;
            }
            else
            {
                tableDesign = TableDesign.None;
            }
        }

        private XElement TblPr => GetOrCreateTablePropertiesSection(); // Xml.Element(DocxNamespace.Main + "tblPr");

        public Alignment Alignment
        {
            get => alignment;
            set
            {
                alignment = value;

                TblPr.Descendants(DocxNamespace.Main + "jc").FirstOrDefault()?.Remove();
                TblPr.Add(new XElement(DocxNamespace.Main + "jc",
                        new XAttribute(DocxNamespace.Main + "val",
                            value.ToString().ToLower())));
            }
        }

        /// <summary>
        /// Auto size this table according to some rule.
        /// </summary>
        public AutoFit AutoFit
        {
            get => autoFit;

            set
            {
                var tableAttributeValue = string.Empty;
                var columnAttributeValue = string.Empty;
                switch (value)
                {
                    case AutoFit.ColumnWidth:
                        {
                            tableAttributeValue = "auto";
                            columnAttributeValue = "dxa";

                            // Disable "Automatically resize to fit contents" option
                            TblPr.GetOrCreateElement(DocxNamespace.Main + "tblLayout")
                                 .GetOrCreateAttribute(DocxNamespace.Main + "type")
                                 .Value = "fixed";
                            break;
                        }

                    case AutoFit.Contents:
                        tableAttributeValue = columnAttributeValue = "auto";
                        break;

                    case AutoFit.Window:
                        tableAttributeValue = columnAttributeValue = "pct";
                        break;

                    case AutoFit.Fixed:
                        {
                            tableAttributeValue = columnAttributeValue = "dxa";

                            var tblLayout = TblPr?.Element(DocxNamespace.Main + "tblLayout");
                            if (tblLayout == null)
                            {
                                var tmp = TblPr?.Element(DocxNamespace.Main + "tblInd") ?? TblPr?.Element(DocxNamespace.Main + "tblW");
                                if (tmp != null)
                                {
                                    tmp.AddAfterSelf(new XElement(DocxNamespace.Main + "tblLayout"));
                                    TblPr.Element(DocxNamespace.Main + "tblLayout")?.SetAttributeValue(DocxNamespace.Main + "type", "fixed");
                                    TblPr.Element(DocxNamespace.Main + "tblW")?.SetAttributeValue(DocxNamespace.Main + "w", ColumnWidths.Sum());
                                }
                                break;
                            }

                            foreach (var type in Xml.LocalNameDescendants("tblLayout").Attributes(DocxNamespace.Main + "type"))
                            {
                                type.Value = "fixed";
                            }

                            TblPr.Element(DocxNamespace.Main + "tblW")?.SetAttributeValue(DocxNamespace.Main + "w", ColumnWidths.Sum());
                            break;
                        }
                }

                // Set preferred width exception type
                foreach (var type in Xml.LocalNameDescendants("tblW").Attributes(DocxNamespace.Main + "type"))
                    type.Value = tableAttributeValue;

                // Set preferred table cell width type
                foreach (var type in Xml.LocalNameDescendants("tcW").Attributes(DocxNamespace.Main + "type"))
                    type.Value = columnAttributeValue;

                autoFit = value;
            }
        }

        /// <summary>
        /// Returns the number of columns in this table.
        /// </summary>
        public int ColumnCount => cachedColCount ?? (RowCount == 0 ? 0 : (cachedColCount = Rows[0].ColumnCount).Value);

        /// <summary>
        /// Gets a list of all column widths for this table.
        /// </summary>
        public double[] ColumnWidths =>
            Xml.Element(DocxNamespace.Main + "tblGrid")?
                .Elements(DocxNamespace.Main + "gridCol")?
                .Select(c => Convert.ToDouble(c.AttributeValue(DocxNamespace.Main + "w")))
                .ToArray();

        /// <summary>
        /// Custom Table Style name
        /// </summary>
        public string CustomTableDesignName
        {
            set
            {
                customTableDesignName = value;
                Design = TableDesign.Custom;
            }

            get => customTableDesignName;
        }

        /// <summary>
        /// The design\style to apply to this table.
        /// </summary>
        public TableDesign Design
        {
            get => tableDesign;
            set
            {
                var style = TblPr.GetOrCreateElement(DocxNamespace.Main + "tblStyle");
                var val = style.GetOrCreateAttribute(DocxNamespace.Main + "val");

                tableDesign = value;

                switch (tableDesign)
                {
                    case TableDesign.None:
                        style?.Remove();
                        break;
                    case TableDesign.Custom when string.IsNullOrEmpty(customTableDesignName):
                        tableDesign = TableDesign.None;
                        style?.Remove();
                        break;
                    case TableDesign.Custom:
                        val.Value = customTableDesignName;
                        break;
                    default:
                        val.Value = tableDesign.GetEnumName();
                        break;
                }

                if (Document.styles == null)
                {
                    var wordStyles = Document.package.GetPart(new Uri("/word/styles.xml", UriKind.Relative));
                    using TextReader tr = new StreamReader(wordStyles.GetStream());
                    Document.styles = XDocument.Load(tr);
                }

                var tableStyle = (
                    from e in Document.styles.Descendants()
                    let styleId = e.Attribute(DocxNamespace.Main + "styleId")
                    where (styleId != null && styleId.Value == val.Value)
                    select e
                ).FirstOrDefault();

                if (tableStyle == null)
                {
                    var externalStyleDoc = Resources.DefaultTableStyles;
                    var styleElement = (
                        from e in externalStyleDoc.Descendants()
                        let styleId = e.Attribute(DocxNamespace.Main + "styleId")
                        where (styleId != null && styleId.Value == val.Value)
                        select e
                    ).First();

                    Document.styles.Element(DocxNamespace.Main + "styles")?.Add(styleElement);
                }
            }
        }

        /// <summary>
        /// Get all of the Hyperlinks in this Table.
        /// </summary>
        public List<Hyperlink> Hyperlinks => Rows.SelectMany(r => r.Hyperlinks).ToList();

        /// <summary>
        /// Returns the index of this Table.
        /// </summary>
        public int Index => Xml.ElementsBeforeSelf().Sum(Paragraph.GetElementTextLength);

        /// <summary>
        /// Returns a list of all Paragraphs inside this container.
        /// </summary>
        ///
        public virtual List<Paragraph> Paragraphs => Rows.SelectMany(r => r.Paragraphs).ToList();

        /// <summary>
        /// Returns a list of all Pictures in a Table.
        /// </summary>
        public List<Picture> Pictures => Rows.SelectMany(r => r.Pictures).ToList();

        /// <summary>
        /// Returns the number of rows in this table.
        /// </summary>
        public int RowCount => Xml.Elements(DocxNamespace.Main + "tr").Count();

        /// <summary>
        /// Returns a list of rows in this table.
        /// </summary>
        public List<Row> Rows => Xml.Elements(DocxNamespace.Main + "tr").Select(r => new Row(this, r)).ToList();

        /// <summary>
        /// Gets or Sets the value of the Table Caption (Alternate Text Title) of this table.
        /// </summary>
        public string TableCaption
        {
            get => TblPr.Element(DocxNamespace.Main + "tblCaption")?.GetVal() ?? string.Empty;
            set
            {
                TblPr?.Descendants(DocxNamespace.Main + "tblCaption").FirstOrDefault()?.Remove();
                TblPr?.Add(new XElement(DocxNamespace.Main + "tblCaption",new XAttribute(DocxNamespace.Main + "val", value)));
            }
        }

        /// <summary>
        /// Gets or Sets the value of the Table Description (Alternate Text Description) of this table.
        /// </summary>
        public string TableDescription
        {
            get => TblPr.Element(DocxNamespace.Main + "tblDescription")?.GetVal() ?? string.Empty;
            set
            {
                TblPr.Descendants(DocxNamespace.Main + "tblDescription").FirstOrDefault()?.Remove();
                TblPr.Add(new XElement(DocxNamespace.Main + "tblDescription", new XAttribute(DocxNamespace.Main + "val", value)));
            }
        }

        /// <summary>
        /// Get a border edge value for this table
        /// </summary>
        /// <param name="borderType">The table border to get</param>
        public Border GetBorder(TableBorderType borderType)
        {
            var border = new Border();
            border.GetDetails(TblPr.Element(DocxNamespace.Main + "tblBorders")?
                                   .Element(DocxNamespace.Main + borderType.GetEnumName()));
            return border;
        }

        /// <summary>
        /// Gets the column width for a given column index.
        /// </summary>
        /// <param name="index"></param>
        public double GetColumnWidth(int index) => ColumnWidths == null || index > ColumnWidths.Length - 1 ? double.NaN : ColumnWidths[index];

        /// <summary>
        /// Insert a column to the right of a Table.
        /// </summary>
        public void InsertColumn()
        {
            InsertColumn(ColumnCount);
        }

        /// <summary>
        /// Insert a column into a table.
        /// </summary>
        /// <param name="index">The index to insert the column at.</param>
        public void InsertColumn(int index)
        {
            if (RowCount > 0)
            {
                cachedColCount = -1;
                foreach (var row in Rows)
                {
                    var cell = HelperFunctions.CreateTableCell();
                    var cells = row.Cells;
                    if (cells.Count == index)
                    {
                        cells[index-1].Xml.AddAfterSelf(cell);
                    }
                    else
                    {
                        cells[index].Xml.AddBeforeSelf(cell);
                    }

                    row.InvalidateCellCache();
                }
            }
        }

        /// <summary>
        /// Insert a row at the end of this table.
        /// </summary>
        public Row InsertRow() => InsertRow(RowCount);

        /// <summary>
        /// Insert a copy of a row at the end of this table.
        /// </summary>
        /// <returns>A new row.</returns>
        public Row InsertRow(Row row) => InsertRow(row, RowCount);

        /// <summary>
        /// Insert a row into this table.
        /// </summary>
        public Row InsertRow(int index)
        {
            if (index < 0 || index > RowCount)
                throw new IndexOutOfRangeException();

            var content = new List<XElement>();
            for (int i = 0; i < ColumnCount; i++)
            {
                double w = 2310d;
                if (columnWidths != null && columnWidths.Length > i)
                    w = columnWidths[i] * 15;

                var cell = HelperFunctions.CreateTableCell(w);
                content.Add(cell);
            }

            return InsertRow(content, index);
        }

        /// <summary>
        /// Insert a copy of a row into this table.
        /// </summary>
        /// <param name="row">Row to copy and insert.</param>
        /// <param name="index">Index to insert row at.</param>
        /// <returns>A new Row</returns>
        public Row InsertRow(Row row, int index)
        {
            if (row == null)
                throw new ArgumentNullException(nameof(row));

            if (index < 0 || index > RowCount)
                throw new IndexOutOfRangeException();

            var content = row.Xml.Elements(DocxNamespace.Main + "tc")
                                              .Select(HelperFunctions.CloneElement).ToList();
            
            return InsertRow(content, index);
        }

        /// <summary>
        /// Merge cells in given column starting with startRow and ending with endRow.
        /// </summary>
        public void MergeCellsInColumn(int columnIndex, int startRow, int endRow)
        {
            // Check for valid start and end indexes.
            if (columnIndex < 0 || columnIndex >= ColumnCount)
                throw new IndexOutOfRangeException(nameof(columnIndex));

            if (startRow < 0 || endRow <= startRow || endRow >= Rows.Count)
                throw new IndexOutOfRangeException();

            for (int rowIndex = startRow; rowIndex < endRow; rowIndex++)
            {
                var c = Rows[rowIndex].Cells[columnIndex];
                _ = c.Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr")
                    .GetOrCreateElement(DocxNamespace.Main + "vMerge");
            }

            var startRowCells = Rows[startRow].Cells;

            var startTcPr = columnIndex > startRowCells.Count 
                ? startRowCells[^1].Xml.Element(DocxNamespace.Main + "tcPr")
                : startRowCells[columnIndex].Xml.Element(DocxNamespace.Main + "tcPr");

            if (startTcPr == null)
            {
                startRowCells[columnIndex].Xml.SetElementValue(DocxNamespace.Main + "tcPr", string.Empty);
                startTcPr = startRowCells[columnIndex].Xml.Element(DocxNamespace.Main + "tcPr");
            }

            startTcPr.GetOrCreateElement(DocxNamespace.Main + "vMerge")
                     .SetAttributeValue(DocxNamespace.Main + "val", "restart");

            for (int rowIndex = startRow; rowIndex < endRow; rowIndex++)
                Rows[rowIndex].InvalidateCellCache();

        }
        /// <summary>
        /// Remove this Table from this document.
        /// </summary>
        public void Remove()
        {
            Xml.Remove();
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
            {
                row.Cells[index].Xml.Remove();
                row.InvalidateCellCache();
            }

            cachedColCount = -1;
        }

        /// <summary>
        /// Remove a row from this Table.
        /// </summary>
        /// <param name="index">The row to remove.</param>
        public void RemoveRow(int index)
        {
            if (index < 0 || index > RowCount - 1)
                throw new IndexOutOfRangeException();

            Rows[index].Xml.Remove();
            if (Rows.Count == 0)
                Remove();
        }

        /// <summary>
        /// Set a table border
        /// </summary>
        /// <param name="borderType">The table border to set</param>
        /// <param name="border">Border object to set the table border</param>
        public void SetBorder(TableBorderType borderType, Border border)
        {
            // Set the border style
            var tblBorders = TblPr.GetOrCreateElement(DocxNamespace.Main + "tblBorders");
            var tblBorderType = tblBorders.GetOrCreateElement(DocxNamespace.Main + borderType.GetEnumName());
            tblBorderType.SetAttributeValue(DocxNamespace.Main + "val", border.Style.GetEnumName());

            // .. and the style
            if (border.Style != BorderStyle.Empty)
            {
                int size = border.Size switch
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
                tblBorderType.SetAttributeValue(DocxNamespace.Main + "sz", size);

                // The space attribute is used for the cell 
                tblBorderType.SetAttributeValue(DocxNamespace.Main + "space", border.SpacingOffset);

                // The color attribute is used for the border color
                tblBorderType.SetAttributeValue(DocxNamespace.Main + "color", border.Color.ToHex());
            }
        }

        /// <summary>
        /// Sets the column width for the given index.
        /// </summary>
        /// <param name="index">Column index</param>
        /// <param name="width">Column width</param>
        public void SetColumnWidth(int index, double width)
        {
            if (ColumnWidths == null || index > ColumnWidths.Length-1)
            {
                if (Rows.Count == 0)
                    throw new Exception("No rows available.");

                columnWidths = Rows[^1].Cells.Select(c => c.Width).ToArray();
            }

            // check if index is matching table columns
            if (index > columnWidths.Length-1)
                throw new Exception($"{nameof(index)} value ({index}) > # of columns ({columnWidths.Length}");

            // get the table grid props
            var grid = Xml.Element(DocxNamespace.Main + "tblGrid");
            if (grid != null)
            {
                grid.RemoveAll();
            }
            else
            {
                grid = new XElement(DocxNamespace.Main + "tblGrid");
                TblPr.AddAfterSelf(grid);
            }

            for (int i = 0; i < columnWidths.Length; i++)
            {
                double value = i == index ? width : columnWidths[i];
                grid.Add(new XElement(DocxNamespace.Main + "gridCol",
                         new XAttribute(DocxNamespace.Main + "w", value)));
            }

            // Reset cell widths
            foreach (var c in Rows.SelectMany(r => r.Cells))
                c.Width = -1;

            // Set to fixed sizing
            AutoFit = AutoFit.Fixed;
        }

        /// <summary>
        /// Set the direction of all content in this Table.
        /// </summary>
        /// <param name="direction">(Left to Right) or (Right to Left)</param>
        public void SetDirection(Direction direction)
        {
            TblPr.Add(new XElement(DocxNamespace.Main + "bidiVisual"));
            Rows.ForEach(r => r.SetDirection(direction));
        }

        /// <summary>
        /// Set the specified cell margin for the table-level.
        /// </summary>
        /// <param name="type">The side of the cell margin.</param>
        /// <param name="margin">The value for the specified cell margin.</param>
        /// <remarks>More information can be found <see cref="http://msdn.microsoft.com/library/documentformat.openxml.wordprocessing.tablecellmargindefault.aspx">here</see></remarks>
        public void SetTableCellMargin(TableCellMarginType type, double margin)
        {
            var tblMargin = TblPr.GetOrCreateElement(DocxNamespace.Main + "tblCellMar")
                                         .GetOrCreateElement(DocxNamespace.Main + type.ToString());
            tblMargin.SetAttributeValue(DocxNamespace.Main + "w", margin);
            tblMargin.SetAttributeValue(DocxNamespace.Main + "type", "dxa");
        }

        /// <summary>
        /// Set the widths for the columns
        /// </summary>
        /// <param name="widths"></param>
        public void SetWidths(double[] widths)
        {
            columnWidths = widths;
            foreach (var row in Rows)
            {
                for (int col = 0; col < widths.Length; col++)
                {
                    var cells = row.Cells.ToList();
                    if (cells.Count > col)
                        cells[col].Width = widths[col];
                }
            }
        }

        /// <summary>
        /// Retrieves or create the table properties (tblPr) section in the document.
        /// </summary>
        /// <returns>The tblPr element for this Table.</returns>
        internal XElement GetOrCreateTablePropertiesSection()
        {
            var tblPrName = DocxNamespace.Main + "tblPr";
            var tblPr = Xml.Element(tblPrName);
            if (tblPr == null)
            {
                Xml.AddFirst(new XElement(tblPrName));
                tblPr = Xml.Element(tblPrName);
            }
            return tblPr;
        }

        private Row InsertRow(IEnumerable<XElement> content, int index)
        {
            var row = new Row(this, new XElement(DocxNamespace.Main + "tr", content));

            XElement rowXml;
            if (index == Rows.Count)
            {
                rowXml = Rows.Last().Xml;
                rowXml.AddAfterSelf(row.Xml);
            }
            else
            {
                rowXml = Rows[index].Xml;
                rowXml.AddBeforeSelf(row.Xml);
            }

            return row;
        }
    }
}