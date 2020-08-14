using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    /// <summary>
    /// Represents a single row in a Table.
    /// </summary>
    public class Row : Container
    {
        /// <summary>
        /// Table owner
        /// </summary>
        internal Table Table { get; set; }

        internal Row(Table table, XElement xml) : base(table.Document, xml)
        {
            Table = table;
            packagePart = table.packagePart;
        }

        /// <summary>
        /// Allow row to break across pages.
        /// The default value is true: Word will break the contents of the row across pages.
        /// If set to false, the contents of the row will not be split across pages, the entire row will be moved to the next page instead.
        /// </summary>
        public bool BreakAcrossPages
        {
            get => Xml.Element(DocxNamespace.Main + "trPr")?
                          .Element(DocxNamespace.Main + "cantSplit") == null;

            set => Xml.GetOrCreateElement(DocxNamespace.Main + "trPr")
                   .SetElementValue(DocxNamespace.Main + "cantSplit", value ? null : string.Empty);
        }

        private ReadOnlyCollection<Cell> cellsCache;

        /// <summary>
        /// A list of Cells in this Row.
        /// </summary>
        public ReadOnlyCollection<Cell> Cells => cellsCache ??= Xml.Elements(DocxNamespace.Main + "tc").Select(c => new Cell(this, c)).ToList().AsReadOnly();

        /// <summary>
        /// Calculates columns count in the row, taking spanned cells into account
        /// </summary>
        public int ColumnCount
        {
            get
            {
                int gridSpanSum = 0, count = 0;
                foreach (var c in Cells)
                {
                    count++;

                    XElement gridSpan = c.Xml.Element(DocxNamespace.Main + "tcPr")?
                                             .Element(DocxNamespace.Main + "gridSpan");
                    XAttribute val = gridSpan?.GetValAttr();
                    if (val != null && int.TryParse(val.Value, out int value))
                    {
                        gridSpanSum += value - 1;
                    }
                }

                return count + gridSpanSum;
            }
        }
        /// <summary>
        /// Height in pixels.
        /// </summary>
        public double Height
        {
            get
            {
                var value = Xml.Element(DocxNamespace.Main + "trPr")?
                               .Element(DocxNamespace.Main + "trHeight")?
                               .GetValAttr();

                if (value == null || !double.TryParse(value.Value, out double heightInWordUnits))
                {
                    value?.Remove();
                    return double.NaN;
                }

                return heightInWordUnits / 15;
            }
            set => SetHeight(value, true);
        }

        public override ReadOnlyCollection<Paragraph> Paragraphs
            => Xml.Descendants(DocxNamespace.Main + "p")
                    .Select(p => new Paragraph(Document, p, 0) { packagePart = Table.packagePart })
                    .ToList()
                    .AsReadOnly();

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
        /// Merge cells starting with startIndex and ending with endIndex.
        /// </summary>
        public void MergeCells(int startIndex, int endIndex)
        {
            // Check for valid start and end indexes.
            if (startIndex < 0 || endIndex <= startIndex || endIndex > Cells.Count + 1)
            {
                throw new IndexOutOfRangeException();
            }

            int gridSpanSum = 0; 
            int value;

            // Foreach each Cell between startIndex and endIndex inclusive.
            foreach (var c in Cells.Where((_, i) => i > startIndex && i <= endIndex).ToList())
            {
                var gsVal = c.Xml.Element(DocxNamespace.Main + "tcPr")?
                                 .Element(DocxNamespace.Main + "gridSpan")?
                                 .GetValAttr();
                if (gsVal != null && int.TryParse(gsVal.Value, out value))
                {
                    gridSpanSum += value - 1;
                }

                // Add this cells Paragraph to the merge start Cell.
                Cells[startIndex].Xml.Add(c.Xml.Elements(DocxNamespace.Main + "p"));

                // Remove this Cell.
                c.Xml.Remove();
            }

            var startTcPr = Cells[startIndex].Xml.GetOrCreateElement(DocxNamespace.Main + "tcPr");
            var startGridSpan = startTcPr.GetOrCreateElement(DocxNamespace.Main + "gridSpan");
            var startVal = startGridSpan.GetValAttr();
            if (startVal != null && int.TryParse(startVal.Value, out value))
            {
                gridSpanSum += value - 1;
            }

            // Set the val attribute to the number of merged cells.
            startGridSpan.SetVal(gridSpanSum + endIndex - startIndex + 1);

            // Reset the cell cache
            cellsCache = null;
        }

        /// <summary>
        /// Remove this cell
        /// </summary>
        public void Remove()
        {
            var tableOwner = Xml.Parent;
            Xml.Remove();
            if (tableOwner?.Elements(DocxNamespace.Main + "tr").Any() == false)
            {
                tableOwner.Remove();
            }

            // Reset the cell cache
            cellsCache = null;
        }

        /// <summary>
        /// Helper method to set either the exact height or the min-height
        /// </summary>
        /// <param name="height">The height value to set (in pixels)</param>
        /// <param name="exact">If true, the height will be forced, otherwise it will be treated as a minimum height, auto growing past it if need be.
        /// </param>
        private void SetHeight(double height, bool exact)
        {
            var trPr = Xml.GetOrCreateElement(DocxNamespace.Main + "trPr");
            var trHeight = trPr.GetOrCreateElement(DocxNamespace.Main + "trHeight");

            // 15 "word units" is equal to one pixel.
            trHeight.SetAttributeValue(DocxNamespace.Main + "hRule", exact ? "exact" : "atLeast");
            trHeight.SetAttributeValue(DocxNamespace.Main + "val", height * 15);
        }

        /// <summary>
        /// Throw away the cell cache
        /// </summary>
        internal void InvalidateCellCache()
        {
            cellsCache = null;
        }
    }
}