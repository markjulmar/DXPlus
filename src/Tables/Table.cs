using DXPlus.Helpers;
using DXPlus.Resources;
using System.IO.Packaging;
using System.Xml.Linq;

namespace DXPlus;

/// <summary>
/// Represents a Table in a document {tbl}
/// </summary>
public class Table : Block, IEquatable<Table>
{
    private string? customTableDesignName;
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
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public Table(int rows, int columns) : this()
    {
        if (rows <= 0) throw new ArgumentOutOfRangeException(nameof(rows), "Rows must be >= 1");
        if (columns <= 0) throw new ArgumentOutOfRangeException(nameof(columns), "Columns must be >= 1");

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
    internal Table(Document? document, PackagePart? packagePart, XElement xml) : base(xml)
    {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        if (xml.Name != Name.Table) throw new ArgumentException($"Root element must be {Name.Table}", nameof(xml));

        if (document != null)
        {
            SetOwner(document, packagePart, false);
        }

        var style = TblPr.Element(Namespace.Main + "tblStyle");
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
        get => ReadTableConditionalFormatting(TblPr);
        set => WriteTableConditionalFormat(TblPr, value);
    }

    /// <summary>
    /// Preferred width for the table
    /// </summary>
    public TableWidth? TableWidth
    {
        get => new(TblPr.Element(Namespace.Main + "tblW"));
        set
        {
            TblPr.Element(Namespace.Main + "tblW")?.Remove();
            if (value == null
                || value.Type == null && value.Width == null) return;
            TblPr.Add(value.Xml);
        }
    }

    /// <summary>
    /// True if the table will auto-fit the contents. This corresponds to the {tblLayout.type} value of the table properties.
    /// </summary>
    public TableLayout? TableLayout
    {
        get => TblPr.Element(Namespace.Main + "tblLayout")?
                    .AttributeValue(Namespace.Main + "type")?
                    .TryGetEnumValue<TableLayout>(out var result) == true ? result : null;

        set => TblPr.GetOrAddElement(Namespace.Main + "tblLayout")
                    .SetAttributeValue(Namespace.Main + "type", value?.GetEnumName());
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
            .SetAttributeValue(Name.MainVal, value.GetEnumName());
    }

    /// <summary>
    /// Indentation in dxa units
    /// </summary>
    public double? Indent
    {
        get
        {
            var value = TblPr.Element(Name.TableIndent)?.Attribute(Namespace.Main + "w");
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
    public IEnumerable<double> DefaultColumnWidths
    {
        get
        {
            var columns = Xml.Element(Namespace.Main + "tblGrid")?
                .Elements(Namespace.Main + "gridCol")
                .Select(c => 
                    double.TryParse(c.AttributeValue(Namespace.Main + "w"), out var dbl) 
                        ? dbl : double.NaN); 
            return columns ?? Enumerable.Empty<double>();
        }
    }

    /// <summary>
    /// Custom Table Style name
    /// </summary>
    public string? CustomTableDesignName
    {
        get => customTableDesignName;
        set
        {
            if (value != null)
            {
                customTableDesignName = value;
                tableDesign = TableDesign.Custom;
                TblPr.GetOrAddElement(Namespace.Main + "tblStyle")
                     .SetAttributeValue(Name.MainVal, value);
            }
            else
            {
                Design = TableDesign.Normal;
            }
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
                throw new ArgumentOutOfRangeException(nameof(Design), $"Cannot set custom design value - use {CustomTableDesignName} property instead.");

            tableDesign = value;

            var style = TblPr.GetOrAddElement(Namespace.Main + "tblStyle");
            if (tableDesign == TableDesign.None)
            {
                style.Remove();
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
    protected override void OnAddToDocument()
    {
        if (ColumnCount == 0 || !Rows.Any())
            throw new Exception("Cannot add empty table to document.");

        // Fix up any unsized columns
        if (DefaultColumnWidths.Any(double.IsNaN))
        {
            OnSetColumnWidths(DefaultColumnWidths.ToArray());
        }

        // Add any required styles
        ApplyTableStyleToDocumentOwner();

        // Set the document/package for each row.
        Rows.ToList().ForEach(r => r.SetOwner(Document, PackagePart, true));
    }

    /// <summary>
    /// This ensures the owning document has the table style applied.
    /// </summary>
    private void ApplyTableStyleToDocumentOwner()
    {
        if (!InDocument) return;

        string? designName = TblPr.Element(Namespace.Main + "tblStyle").GetVal();
        if (string.IsNullOrWhiteSpace(designName) 
            || string.Compare(designName, "none", StringComparison.InvariantCultureIgnoreCase)==0)
            return;

        if (!Document.Styles.HasStyle(designName, StyleType.Table))
        {
            var styleElement = Resource.DefaultTableStyles()
                .Descendants()
                .FindByAttrVal(Namespace.Main + "styleId", designName);
            Document.Styles.Add(styleElement!);
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
            TblPr.Descendants(Namespace.Main + "tblCaption").FirstOrDefault()?.Remove();
            TblPr.Add(new XElement(Namespace.Main + "tblCaption", new XAttribute(Name.MainVal, value)));
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
    public void AddColumn() => InsertColumn(ColumnCount);

    /// <summary>
    /// Insert a column into a table. All rows are given an empty cell.
    /// </summary>
    /// <param name="index">The index to insert the column at.</param>
    public void InsertColumn(int index)
    {
        if (index < 0 || index > ColumnCount)
            throw new ArgumentOutOfRangeException(nameof(index));

        bool insertAtEnd = (index == ColumnCount);

        // Create a new column by splitting the last column in half.
        var grid = Xml.GetOrAddElement(Namespace.Main + "tblGrid");
        var gridColumns = grid.Elements(Namespace.Main + "gridCol").ToList();

        XElement existingGridCol = insertAtEnd ? gridColumns.Last() : gridColumns[index];
        double.TryParse(existingGridCol.AttributeValue(Namespace.Main + "w"), out double width);
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
    public TableRow AddRow() => InsertRow(Rows.Count());

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
                    cell.Text = string.Empty;
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
        Xml.Element(Namespace.Main + "tblGrid")!.Elements(Namespace.Main + "gridCol").ElementAt(index).Remove();

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
    /// Table border
    /// </summary>
    private static readonly XName tblBorders = Namespace.Main + "tblBorders";

    /// <summary>
    /// Top border for this table
    /// </summary>
    public Border? TopBorder
    {
        get => new(BorderType.Top, TblPr, tblBorders);
        set => Border.SetElementValue(BorderType.Top, TblPr, tblBorders, value);
    }

    /// <summary>
    /// Bottom border for this table
    /// </summary>
    public Border? BottomBorder
    {
        get => new(BorderType.Bottom, TblPr, tblBorders);
        set => Border.SetElementValue(BorderType.Bottom, TblPr, tblBorders, value);
    }

    /// <summary>
    /// Left border for this table
    /// </summary>
    public Border? LeftBorder
    {
        get => new(BorderType.Left, TblPr, tblBorders);
        set => Border.SetElementValue(BorderType.Left, TblPr, tblBorders, value);
    }

    /// <summary>
    /// Right border for this table
    /// </summary>
    public Border? RightBorder
    {
        get => new(BorderType.Right, TblPr, tblBorders);
        set => Border.SetElementValue(BorderType.Right, TblPr, tblBorders, value);
    }

    /// <summary>
    /// Inside horizontal border for this table
    /// </summary>
    public Border? InsideHorizontalBorder
    {
        get => new(BorderType.InsideH, TblPr, tblBorders);
        set => Border.SetElementValue(BorderType.InsideH, TblPr, tblBorders, value);
    }

    /// <summary>
    /// Inside horizontal border for this table
    /// </summary>
    public Border? InsideVerticalBorder
    {
        get => new(BorderType.InsideV, TblPr, tblBorders);
        set => Border.SetElementValue(BorderType.InsideV, TblPr, tblBorders, value);
    }

    /// <summary>
    /// Supply all the column widths
    /// </summary>
    /// <param name="widths"></param>
    public void SetColumnWidths(IEnumerable<double> widths)
    {
        var w = widths.ToArray();
        if (w.Length != ColumnCount)
            throw new ArgumentOutOfRangeException(nameof(widths), "Must supply widths for each column.");

        OnSetColumnWidths(w);
    }

    /// <summary>
    /// Sets the column width for the given index.
    /// </summary>
    /// <param name="index">Column index</param>
    /// <param name="width">Column width</param>
    public void SetColumnWidth(int index, double width)
    {
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        if (width < 0) throw new ArgumentOutOfRangeException(nameof(width));

        double[] columnWidths = DefaultColumnWidths.ToArray();
        if (index > columnWidths.Length - 1)
        {
            if (!Rows.Any())
            {
                throw new InvalidOperationException("Must have at least one row to determine column widths.");
            }

            columnWidths = Rows.ToList()[^1].Cells.Select(c => c.CellWidth?.Width ?? 0).ToArray();
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
        if (columnWidths.Any(double.IsNaN))
        {
            double totalSpace;
            if (TableWidth?.Type == TableWidthUnit.Dxa || !InDocument)
                totalSpace = TableWidth?.Width ?? PageSize.LetterWidth;
            else
            {
                totalSpace = Document.Sections.First().Properties.AdjustedPageWidth;
                if (TableWidth?.Type == TableWidthUnit.Percentage)
                {
                    totalSpace *= (TableWidth?.Width ?? 100*TableWidth.PctMultiplier)/ TableWidth.PctMultiplier;
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
        var grid = Xml.Element(Namespace.Main + "tblGrid");
        if (grid != null)
        {
            grid.RemoveAll();
        }
        else
        {
            grid = new XElement(Namespace.Main + "tblGrid");
            TblPr.AddAfterSelf(grid);
        }

        foreach (var width in columnWidths)
        {
            grid.Add(new XElement(Namespace.Main + "gridCol",
                new XAttribute(Namespace.Main + "w", width)));
        }

        // Reset cell widths
        foreach (var row in Rows)
        {
            row.SetColumnWidths(columnWidths);
        }
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
            var cellMargin = TblPr.GetOrAddElement(Namespace.Main + "tblCellMar")
                .GetOrAddElement(Namespace.Main + type.GetEnumName());
            cellMargin.SetAttributeValue(Namespace.Main + "w", margin);
            cellMargin.SetAttributeValue(Namespace.Main + "type", "dxa");
        }
        else
        {
            var margins = TblPr.Element(Namespace.Main + "tblCellMar");
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
    /// Read the hex value for table conditional formatting and turn it back
    /// into an enumeration.
    /// </summary>
    /// <param name="tblPr">Table properties</param>
    /// <returns>Enum value</returns>
    private static TableConditionalFormatting ReadTableConditionalFormatting(XElement tblPr)
    {
        string? value = tblPr.Element(Namespace.Main + "tblLook")?.GetVal();
        if (!string.IsNullOrEmpty(value))
        {
            // It's represented as a hex value, we need to turn it back to
            // an integer base-10 value to use Enum.Parse.
            if (int.TryParse(value, System.Globalization.NumberStyles.HexNumber,
                    null, out int num))
            {
                return Enum.TryParse<TableConditionalFormatting>(
                    num.ToString(), out var tcf)
                    ? tcf
                    : TableConditionalFormatting.None;
            }
        }
        return TableConditionalFormatting.None;
    }

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
    private static XElement CreateEmptyTableXml() => new(Name.Table,
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
                    DocumentHelpers.GenerateHexId()))
        );
    }

    /// <summary>
    /// Determines equality for a table object
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Table? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a table
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Table);

    /// <summary>
    /// Returns hashcode for this table
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();
}