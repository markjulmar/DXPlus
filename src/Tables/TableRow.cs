using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a single row in a Table.
    /// </summary>
    public class TableRow : DocXElement, IEquatable<TableRow>
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
        internal TableRow(Table table, XElement xml) : base(table.Document, table.PackagePart, xml)
        {
            Table = table;
        }

        /// <summary>
        /// Retrieve the row properites. This is optional and can be null.
        /// </summary>
        private XElement trPr(bool create = false) => create ? Xml.GetOrAddElement(Namespace.Main + "trPr") : Xml.Element(Namespace.Main + "trPr");

        /// <summary>
        /// Allow row to break across pages.
        /// The default value is true: Word will break the contents of the row across pages.
        /// If set to false, the contents of the row will not be split across pages, the
        /// entire row will be moved to the next page instead.
        /// </summary>
        public bool BreakAcrossPages
        {
            get => trPr(false)?.Element(Namespace.Main + "cantSplit") == null;

            set => trPr(true).SetElementValue(Namespace.Main + "cantSplit", value ? null : string.Empty);
        }

        /// <summary>
        /// A list of Cells in this Row.
        /// </summary>
        public IReadOnlyList<TableCell> Cells => Xml.Elements(Namespace.Main + "tc").Select(e => new TableCell(this, e)).ToList();

        /// <summary>
        /// Calculates columns count in the row, taking spanned cells into account
        /// </summary>
        public int ColumnCount
        {
            get
            {
                IReadOnlyList<TableCell> cells = Cells;
                return cells.Count + cells.Select(cell => cell.GridSpan - 1).Sum();
            }
        }

        /// <summary>
        /// Row height (can be null). Value is in dxa units
        /// </summary>
        public double? Height
        {
            get
            {
                XAttribute value = trPr()?.Element(Namespace.Main + "trHeight")?.GetValAttr();
                if (value == null || !double.TryParse(value.Value, out double heightInDxa))
                {
                    value?.Remove();
                    return null;
                }

                return heightInDxa;
            }
            set => SetHeight(value, true);
        }

        /// <summary>
        /// Set to true to make this row the table header row that will be repeated on each page
        /// </summary>
        public bool TableHeader
        {
            get => trPr(false)?.Element(Namespace.Main + "tblHeader") != null;
            set
            {
                var trPr = this.trPr(true);
                XElement tblHeader = trPr.Element(Namespace.Main + "tblHeader");

                if (tblHeader == null && value)
                {
                    trPr.SetElementValue(Namespace.Main + "tblHeader", string.Empty);
                }
                else if (tblHeader != null && !value)
                {
                    tblHeader.Remove();
                }
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
            {
                throw new IndexOutOfRangeException(nameof(startIndex));
            }

            if (endIndex >= Cells.Count)
            {
                throw new IndexOutOfRangeException(nameof(count));
            }

            IReadOnlyList<TableCell> cells = Cells;
            TableCell startCell = cells[startIndex];
            int gridSpanSum = 0;

            // Merge all the cells beyond startIndex up to the ending index.
            for (int i = startIndex; i <= endIndex; i++)
            {
                TableCell cell = cells[i];
                gridSpanSum += cell.GridSpan - 1;

                // Add the contents of the cell to the starting cell and remove it.
                if (cell != startCell)
                {
                    startCell.Xml.Add(cell.Xml.Elements(Name.Paragraph).Where(p => !p.IsEmpty));
                    cell.Xml.Remove();

                    // Cells were all empty -- just add one empty paragraph.
                    if (startCell.Xml.Element(Name.Paragraph) == null)
                    {
                        startCell.Xml.Add(new XElement(Name.Paragraph,
                        new XAttribute(Name.ParagraphId,
                            HelperFunctions.GenerateHexId())));
                    }
                }
            }

            // Set the gridSpan to the number of merged cells.
            startCell.Xml.GetOrAddElement(Namespace.Main + "tcPr")
                         .GetOrAddElement(Namespace.Main + "gridSpan")
                         .SetAttributeValue(Name.MainVal, gridSpanSum + endIndex - startIndex + 1);
        }

        /// <summary>
        /// Remove this row. If this is the last row, an exception is thrown.
        /// </summary>
        public void Remove()
        {
            if (Table.Rows.Count() == 1)
                throw new Exception("Cannot remove final row from table. Must delete table instead.");

            Xml.Remove();
        }

        /// <summary>
        /// Helper method to set either the exact height or the min-height
        /// </summary>
        /// <param name="height">The height value to set in dxa units</param>
        /// <param name="exact">If true, the height will be forced, otherwise it will be treated as a minimum height, auto growing past it if need be.
        /// </param>
        private void SetHeight(double? height, bool exact)
        {
            if (height != null)
            {
                XElement trHeight = trPr(true).GetOrAddElement(Namespace.Main + "trHeight");
                trHeight.SetAttributeValue(Namespace.Main + "hRule", exact ? "exact" : "atLeast");
                trHeight.SetAttributeValue(Name.MainVal, height);
            }
            else
            {
                trPr()?.Element(Namespace.Main + "trHeight")?.Remove();
            }
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
                if (!cells.Any(c => c.GridSpan > 0))
                    throw new Exception($"Row column count {cells.Count} does not match passed width count {columnWidths.Length}.");

                // The passed array can have more values that we
                // have columns if there are merged cells in this row.
                // In this case, sum the columns leading up to the merge
                List<double> cw = new(); int cellIndex = 0;
                for (int index = 0; index < columnWidths.Length; index++)
                {
                    int span = cells[cellIndex++].GridSpan;
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
                cells[index].SetWidth(TableWidthUnit.Dxa, columnWidths[index]);
            }
        }

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected override void OnDocumentOwnerChanged()
        {
            foreach (TableCell cell in Cells)
                cell.SetOwner(Document, PackagePart);
        }

        /// <summary>
        /// Determines equality for table rows
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(TableRow other)
        {
            if (other == null)
                return false;
            if (ReferenceEquals(this, other))
                return true;
            return Xml == other.Xml;
        }
    }
}