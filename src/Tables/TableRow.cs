using System.Diagnostics;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a single row in a Table.
/// </summary>
public sealed class TableRow : DocXElement, IEquatable<TableRow>
{
    /// <summary>
    /// Table owner
    /// </summary>
    public Table Table { get; }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="table"></param>
    /// <param name="xml"></param>
    internal TableRow(Table table, XElement xml) : base(xml)
    {
        Table = table;
        if (table.InDocument)
        {
            SetOwner(table.Document, table.PackagePart, false);
        }
    }

    /// <summary>
    /// Table row properties
    /// </summary>
    public TableRowProperties Properties => new(Xml.GetOrInsertElement(Name.TableRowProperties));

    /// <summary>
    /// A list of Cells in this Row.
    /// </summary>
    public IReadOnlyList<TableCell> Cells 
        => Xml.Elements(Namespace.Main + "tc")
              .Select(e => new TableCell(this, e))
              .ToList();

    /// <summary>
    /// Calculates columns count in the row, taking spanned cells into account
    /// </summary>
    public int ColumnCount
    {
        get
        {
            var cells = Cells;
            return cells.Count + cells.Select(cell => cell.Properties.GridSpan - 1).Sum();
        }
    }

    /// <summary>
    /// Merge cells starting with startIndex up to count
    /// </summary>
    public void MergeCells(int startIndex, int count)
    {
        int endIndex = startIndex + count - 1;

        // Check for valid start and end indexes.
        if (startIndex < 0 || startIndex >= endIndex)
            throw new IndexOutOfRangeException(nameof(startIndex));

        if (endIndex >= Cells.Count)
            throw new IndexOutOfRangeException(nameof(count));

        IReadOnlyList<TableCell> cells = Cells;
        TableCell startCell = cells[startIndex];
        int gridSpanSum = 0;

        // Merge all the cells beyond startIndex up to the ending index.
        for (int i = startIndex; i <= endIndex; i++)
        {
            var cell = cells[i];
            gridSpanSum += cell.Properties.GridSpan - 1;

            // Add the contents of the cell to the starting cell and remove it.
            if (!ReferenceEquals(cell, startCell))
            {
                startCell.Xml.Add(cell.Xml.Elements(Name.Paragraph).Where(p => !p.IsEmpty));
                cell.Xml.Remove();

                // Cells were all empty -- just add one empty paragraph.
                if (startCell.Xml.Element(Name.Paragraph) == null)
                {
                    startCell.Xml.Add(new XElement(Name.Paragraph,
                        new XAttribute(Name.ParagraphId,
                            DocumentHelpers.GenerateHexId())));
                }
            }
        }

        // Set the gridSpan to the number of merged cells.
        startCell.Xml.GetOrAddElement(Name.TableCellProperties)
            .GetOrAddElement(Namespace.Main + "gridSpan")
            .SetAttributeValue(Name.MainVal, gridSpanSum + endIndex - startIndex + 1);
    }

    /// <summary>
    /// Remove this row. If this is the last row, an exception is thrown.
    /// </summary>
    public void Remove()
    {
        if (Table.Rows.Count == 1)
            throw new Exception("Cannot remove final row from table. Must delete table instead.");

        Xml.Remove();
    }

    /// <summary>
    /// Reset the column widths for this row. This is called
    /// when the table changes shape.
    /// </summary>
    /// <param name="columnWidths">Array of column widths</param>
    internal void SetColumnWidths(double[] columnWidths)
    {
        var cells = Cells;
        if (cells.Count != columnWidths.Length)
        {
            if (!cells.Any(c => c.Properties.GridSpan > 0))
                throw new Exception($"Row column count {cells.Count} does not match passed width count {columnWidths.Length}.");

            // The passed array can have more values that we
            // have columns if there are merged cells in this row.
            // In this case, sum the columns leading up to the merge
            List<double> cw = new(); int cellIndex = 0;
            for (int index = 0; index < columnWidths.Length; index++)
            {
                int span = cells[cellIndex++].Properties.GridSpan;
                double val = columnWidths[index];
                if (span == 1)
                    cw.Add(val);
                else
                {
                    for (int x = 1; x < span; x++)
                    {
                        index++;
                        val += columnWidths[index];
                    }
                    cw.Add(val);
                }
            }

            columnWidths = cw.ToArray();
            Debug.Assert(cells.Count == columnWidths.Length);
        }

        // Assign the widths
        for (int index = 0; index < cells.Count; index++)
        {
            cells[index].Properties.CellWidth = new TableElementWidth(TableWidthUnit.Dxa, columnWidths[index]);
        }
    }

    /// <summary>
    /// Called when the document owner is changed.
    /// </summary>
    protected override void OnAddToDocument()
    {
        foreach (var cell in Cells)
        {
            cell.SetOwner(Document, PackagePart, true);
        }
    }

    /// <summary>
    /// Determines equality for table rows
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(TableRow? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);
}