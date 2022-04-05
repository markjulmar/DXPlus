using DXPlus.Resources;
using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a Table in a document {tbl}
/// </summary>
public sealed class Table : Block, IEquatable<Table>
{
    private List<TableRow>? rowCache;

    /// <summary>
    /// Public constructor to create an empty table
    /// </summary>
    public Table() : this(null, null, CreateEmptyTableXml())
    {
    }

    /// <summary>
    /// Create a table from a multidimensional array of strings.
    /// </summary>
    /// <param name="values"></param>
    public Table(string[,] values) : this(values.GetLength(0), values.GetLength(1))
    {
        if (!values.IsFixedSize)
            throw new ArgumentOutOfRangeException(nameof(values), "Values must be a fixed multi-dimensional array.");

        int totalRows = values.GetLength(0);
        int totalColumns = values.GetLength(1);

        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++)
        {
            var row = Rows[rowIndex];
            for (int colIndex = 0; colIndex < totalColumns; colIndex++)
            {
                var cell = row.Cells[colIndex];
                cell.Text = values[rowIndex, colIndex];
            }
        }
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

        Properties.ConditionalFormatting = TableConditionalFormatting.None;

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
            SetOwner(document, packagePart, false);
    }

    /// <summary>
    /// Properties applied to this table
    /// </summary>
    public TableProperties Properties
    {
        get => new(Xml.GetOrInsertElement(Name.TableProperties), this);

        set
        {
            var tPr = Xml.Element(Name.TableProperties);
            tPr?.Remove();

            var xml = value.Xml!;
            if (xml.Parent != null)
                xml = xml.Clone();

            Xml.AddFirst(xml);

            if (InDocument) 
                EnsureTableStyleInDocument();
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
    /// Called when the document owner is changed.
    /// </summary>
    protected override void OnAddToDocument()
    {
        // Force a regen of the rows.
        rowCache = null;

        if (ColumnCount == 0 || Rows.Count == 0)
            throw new Exception("Cannot add empty table to document.");

        // Fix up any unsized columns
        if (DefaultColumnWidths.Any(double.IsNaN))
        {
            OnSetColumnWidths(DefaultColumnWidths.ToArray());
        }

        // Add any required styles
        EnsureTableStyleInDocument();

        // Set the document/package for each row.
        Rows.ToList().ForEach(r => r.SetOwner(Document, PackagePart, true));
    }

    /// <summary>
    /// This ensures the owning document has the table style applied.
    /// </summary>
    internal void EnsureTableStyleInDocument()
    {
        string? designName = Properties.Design;
        if (string.IsNullOrWhiteSpace(designName)) return;

        if (!Document.Styles.Exists(designName, StyleType.Table))
        {
            var styleElement = Resource.DefaultTableStyles()
                .Descendants()
                .FindByAttrVal(Namespace.Main + "styleId", designName);
            
            // if it's not a default style we know about, then ignore it.
            if (styleElement == null) 
                return;
            
            styleElement.SetAttributeValue(Namespace.Main + "type", StyleType.Table.GetEnumName());
            Document.Styles.AddStyle(styleElement);
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
    public IEnumerable<Drawing> Drawings => Rows.SelectMany(r => r.Cells).SelectMany(c => c.Drawings);

    /// <summary>
    /// Returns a list of rows in this table.
    /// </summary>
    public IReadOnlyList<TableRow> Rows
    {
        get
        {
            return rowCache ??= 
                new List<TableRow>(Xml.Elements(Namespace.Main + "tr").Select(r => new TableRow(this, r)));
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

        rowCache = null;
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
    public TableRow AddRow() => InsertRow(Rows.Count);

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

        rowCache = null;
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

        if (startRow < 0 || startRow > endRow)
            throw new IndexOutOfRangeException(nameof(startRow));

        if (endRow > Rows.Count)
            throw new IndexOutOfRangeException(nameof(count));

        var startRowElement = Rows.ElementAt(startRow).Cells[columnIndex].Xml;

        // Move the content over and add vMerge to each row cell
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
        {
            var cell = Rows[rowIndex].Cells[columnIndex];
            var vMerge = cell.Xml.GetOrAddElement(Namespace.Main + "tcPr")
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
        var rows = Rows.ToList();
        if (index < 0 || index > rows.Count - 1)
        {
            throw new IndexOutOfRangeException();
        }

        rows[index].Xml.Remove();
        rowCache = null;
        if (!Rows.Any()) // use real property
            Remove();
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

            columnWidths = Rows.ToList()[^1].Cells.Select(c => c.Properties.CellWidth?.Width ?? 0).ToArray();
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
            var tableWidth = Properties.TableWidth;

            double totalSpace;
            if (tableWidth?.Type == TableWidthUnit.Dxa || !InDocument)
            {
                totalSpace = tableWidth?.Width ?? 0;
                if (totalSpace == 0)
                    totalSpace = PageSize.LetterWidth;
            }
            else
            {
                totalSpace = Document.Sections.First().Properties.AdjustedPageWidth;
                if (tableWidth?.Type == TableWidthUnit.Percentage)
                {
                    totalSpace *= (tableWidth.Width ?? 100*TableElementWidth.PctMultiplier)/ TableElementWidth.PctMultiplier;
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
            Properties.Xml!.AddAfterSelf(grid);
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
    /// Create and return the default {tbl} root element
    /// </summary>
    private static XElement CreateEmptyTableXml() => new(Name.Table,
        TableProperties.CreateDefaultTableProperties(),
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