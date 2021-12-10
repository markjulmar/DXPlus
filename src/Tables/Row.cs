using DXPlus.Helpers;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Represents a single row in a Table.
    /// </summary>
    public class Row : DocXElement
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
        internal Row(Table table, XElement xml) : base(table.Document, xml)
        {
            Table = table;
        }

        /// <summary>
        /// Allow row to break across pages.
        /// The default value is true: Word will break the contents of the row across pages.
        /// If set to false, the contents of the row will not be split across pages, the
        /// entire row will be moved to the next page instead.
        /// </summary>
        public bool BreakAcrossPages
        {
            get => Xml.Element(Namespace.Main + "trPr")?
                      .Element(Namespace.Main + "cantSplit") == null;

            set => Xml.GetOrAddElement(Namespace.Main + "trPr")
                      .SetElementValue(Namespace.Main + "cantSplit", value ? null : string.Empty);
        }

        /// <summary>
        /// A list of Cells in this Row.
        /// </summary>
        public IReadOnlyList<Cell> Cells => Xml.Elements(Namespace.Main + "tc").Select(e => new Cell(this, e)).ToList();

        /// <summary>
        /// Calculates columns count in the row, taking spanned cells into account
        /// </summary>
        public int ColumnCount
        {
            get
            {
                IReadOnlyList<Cell> cells = Cells;
                return cells.Count + cells.Select(cell => cell.GridSpan - 1).Sum();
            }
        }

        /// <summary>
        /// Height in pixels.
        /// </summary>
        public double? Height
        {
            get
            {
                XAttribute value = Xml.Element(Namespace.Main + "trPr")?
                               .Element(Namespace.Main + "trHeight")?
                               .GetValAttr();

                if (value == null || !double.TryParse(value.Value, out double heightInWordUnits))
                {
                    value?.Remove();
                    return null;
                }

                return heightInWordUnits / TableHelpers.UnitConversion;
            }
            set => SetHeight(value, true);
        }

        /// <summary>
        /// Set to true to make this row the table header row that will be repeated on each page
        /// </summary>
        public bool TableHeader
        {
            get => Xml.Element(Namespace.Main + "trPr")?
                      .Element(Namespace.Main + "tblHeader") != null;
            set
            {
                XElement trPr = Xml.GetOrAddElement(Namespace.Main + "trPr");
                XElement tblHeader = trPr.Element(Namespace.Main + "tblHeader");

                if (tblHeader == null && value)
                {
                    trPr.SetElementValue(Namespace.Main + "tblHeader", string.Empty);
                }
                if (tblHeader != null && !value)
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

            IReadOnlyList<Cell> cells = Cells;
            Cell startCell = cells[startIndex];
            int gridSpanSum = 0;

            // Merge all the cells beyond startIndex up to the ending index.
            for (int i = startIndex; i <= endIndex; i++)
            {
                Cell cell = cells[i];
                gridSpanSum += cell.GridSpan - 1;

                // Add the contents of the cell to the starting cell and remove it.
                if (cell != startCell)
                {
                    startCell.Xml.Add(cell.Xml.Elements(Name.Paragraph));
                    cell.Xml.Remove();
                }
            }

            // Set the gridSpan to the number of merged cells.
            startCell.Xml.GetOrAddElement(Namespace.Main + "tcPr")
                         .GetOrAddElement(Namespace.Main + "gridSpan")
                         .SetAttributeValue(Name.MainVal, gridSpanSum + endIndex - startIndex + 1);
        }

        /// <summary>
        /// Remove this row
        /// </summary>
        public void Remove()
        {
            XElement tableOwner = Xml.Parent;
            Xml.Remove();
            if (tableOwner?.Elements(Namespace.Main + "tr").Any() == false)
            {
                tableOwner.Remove();
            }
        }

        /// <summary>
        /// Helper method to set either the exact height or the min-height
        /// </summary>
        /// <param name="height">The height value to set (in pixels)</param>
        /// <param name="exact">If true, the height will be forced, otherwise it will be treated as a minimum height, auto growing past it if need be.
        /// </param>
        private void SetHeight(double? height, bool exact)
        {
            if (height != null)
            {
                XElement trPr = Xml.GetOrAddElement(Namespace.Main + "trPr");
                XElement trHeight = trPr.GetOrAddElement(Namespace.Main + "trHeight");
                trHeight.SetAttributeValue(Namespace.Main + "hRule", exact ? "exact" : "atLeast");
                trHeight.SetAttributeValue(Name.MainVal, height * TableHelpers.UnitConversion);
            }
            else
            {
                Xml.Element(Namespace.Main + "trPr")?
                    .Element(Namespace.Main + "trHeight")?
                    .Remove();
            }
        }

        /// <summary>
        /// Reset the column widths - used when the table changes shape.
        /// </summary>
        /// <param name="columnWidths">New column widths</param>
        internal void SetColumnWidths(IList<double> columnWidths)
        {
            IReadOnlyList<Cell> cells = Cells;
            if (cells.Count != columnWidths.Count)
            {
                throw new Exception($"Row column count {cells.Count} doesn't match passed column count {columnWidths.Count}.");
            }

            int index = 0;
            foreach (Cell cell in cells)
            {
                cell.Width = columnWidths[index++];
            }
        }

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected override void OnDocumentOwnerChanged(IDocument previousValue, IDocument newValue)
        {
            base.OnDocumentOwnerChanged(previousValue, newValue);
            foreach (Cell cell in Cells)
            {
                cell.Document = (DocX)newValue;
            }
        }

        /// <summary>
        /// Called when the package part is changed.
        /// </summary>
        protected override void OnPackagePartChanged(PackagePart previousValue, PackagePart newValue)
        {
            base.OnPackagePartChanged(previousValue, newValue);
            foreach (Cell cell in Cells)
            {
                cell.PackagePart = newValue;
            }
        }
    }
}