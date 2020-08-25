using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

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
            PackagePart = table.PackagePart;
        }

        /// <summary>
        /// Allow row to break across pages.
        /// The default value is true: Word will break the contents of the row across pages.
        /// If set to false, the contents of the row will not be split across pages, the
        /// entire row will be moved to the next page instead.
        /// </summary>
        public bool BreakAcrossPages
        {
            get => Xml.Element(DocxNamespace.Main + "trPr")?
                      .Element(DocxNamespace.Main + "cantSplit") == null;

            set => Xml.GetOrCreateElement(DocxNamespace.Main + "trPr")
                      .SetElementValue(DocxNamespace.Main + "cantSplit", value ? null : string.Empty);
        }

        /// <summary>
        /// A list of Cells in this Row.
        /// </summary>
        public IReadOnlyList<Cell> Cells => Xml.Elements(DocxNamespace.Main + "tc").Select(e => new Cell(this, e)).ToList();

        /// <summary>
        /// Calculates columns count in the row, taking spanned cells into account
        /// </summary>
        public int ColumnCount
        {
            get
            {
                var cells = Cells;
                return cells.Count + cells.Select(cell => cell.GridSpan-1).Sum();
            }
        }

        /// <summary>
        /// Height in pixels.
        /// </summary>
        public double? Height
        {
            get
            {
                var value = Xml.Element(DocxNamespace.Main + "trPr")?
                               .Element(DocxNamespace.Main + "trHeight")?
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
            get => Xml.Element(DocxNamespace.Main + "trPr")?
                      .Element(DocxNamespace.Main + "tblHeader") != null;
            set
            {
                var trPr = Xml.GetOrCreateElement(DocxNamespace.Main + "trPr");
                var tblHeader = trPr.Element(DocxNamespace.Main + "tblHeader");

                if (tblHeader == null && value)
                {
                    trPr.SetElementValue(DocxNamespace.Main + "tblHeader", string.Empty);
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
                throw new IndexOutOfRangeException(nameof(startIndex));
            if (endIndex >= Cells.Count)
                throw new IndexOutOfRangeException(nameof(count));

            var cells = Cells;
            var startCell = cells[startIndex];
            int gridSpanSum = 0;

            // Merge all the cells beyond startIndex up to the ending index.
            for (int i = startIndex; i <= endIndex; i++)
            {
                var cell = cells[i];
                gridSpanSum += cell.GridSpan-1;

                // Add the contents of the cell to the starting cell and remove it.
                if (cell != startCell)
                {
                    startCell.Xml.Add(cell.Xml.Elements(DocxNamespace.Main + "p"));
                    cell.Xml.Remove();
                }
            }

            // Set the gridSpan to the number of merged cells.
            startCell.Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr")
                         .GetOrCreateElement(DocxNamespace.Main + "gridSpan")
                         .SetAttributeValue(DocxNamespace.Main + "val", gridSpanSum + endIndex - startIndex + 1);
        }

        /// <summary>
        /// Remove this row
        /// </summary>
        public void Remove()
        {
            var tableOwner = Xml.Parent;
            Xml.Remove();
            if (tableOwner?.Elements(DocxNamespace.Main + "tr").Any() == false)
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
                var trPr = Xml.GetOrCreateElement(DocxNamespace.Main + "trPr");
                var trHeight = trPr.GetOrCreateElement(DocxNamespace.Main + "trHeight");
                trHeight.SetAttributeValue(DocxNamespace.Main + "hRule", exact ? "exact" : "atLeast");
                trHeight.SetAttributeValue(DocxNamespace.Main + "val", height * TableHelpers.UnitConversion);
            }
            else
            {
                Xml.Element(DocxNamespace.Main + "trPr")?
                    .Element(DocxNamespace.Main + "trHeight")?
                    .Remove();

            }

        }

        /// <summary>
        /// Reset the column widths - used when the table changes shape.
        /// </summary>
        /// <param name="columnWidths">New column widths</param>
        internal void SetColumnWidths(IList<double> columnWidths)
        {
            var cells = Cells;
            if (cells.Count != columnWidths.Count)
                throw new Exception($"Row column count {cells.Count} doesn't match passed column count {columnWidths.Count}.");

            int index = 0;
            foreach (var cell in cells)
                cell.Width = columnWidths[index++];
        }

        /// <summary>
        /// Called when the document owner is changed.
        /// </summary>
        protected override void OnDocumentOwnerChanged(DocX previousValue, DocX newValue)
        {
            base.OnDocumentOwnerChanged(previousValue, newValue);

            PackagePart = Table?.PackagePart;
            foreach (var cell in Cells)
                cell.Document = newValue;
        }
    }
}