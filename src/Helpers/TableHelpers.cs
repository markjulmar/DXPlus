using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    /// <summary>
    /// Helper methods to create/manage table (w:tbl) elements
    /// </summary>
    internal static class TableHelpers
    {
        public const double UnitConversion = 20.0;  // Height/Widths are represented in 20ths of a pt.
        public const int MaxTableWidth = (int) (9350.0 / UnitConversion);

        public static XElement CreateTable(int rows, int columns)
        {
            if (rows < 1)
                throw new ArgumentOutOfRangeException(nameof(rows), "Cannot be < 1");
            if (columns < 1)
                throw new ArgumentOutOfRangeException(nameof(columns), "Cannot be < 1");

            var columnWidths = CalculateProportionalWidths(columns, MaxTableWidth);
            return CreateTable(rows, columnWidths);
        }

        /// <summary>
        /// Creates an array of columns widths for a new table.
        /// </summary>
        /// <param name="columnCount"></param>
        /// <param name="maxSize"></param>
        /// <returns></returns>
        public static int[] CalculateProportionalWidths(int columnCount, int maxSize)
        {
            int[] columnWidths = new int[columnCount];
            Array.Fill(columnWidths, maxSize / columnCount);
            return columnWidths;
        }

        /// <summary>
        /// Write the element children for the TableConditionalFormatting
        /// </summary>
        /// <param name="tblPr"></param>
        /// <param name="format"></param>
        public static void WriteTableConditionalFormat(XElement tblPr, TableConditionalFormatting format)
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
        /// Method to generate the XML for a basic table
        /// </summary>
        /// <param name="rows">Number of rows</param>
        /// <param name="columnWidths">Column widths</param>
        /// <returns></returns>
        private static XElement CreateTable(int rows, int[] columnWidths)
        {
            if (columnWidths.Sum() > MaxTableWidth)
                throw new ArgumentOutOfRangeException(nameof(columnWidths), $"columnWidths sum cannot exceed {MaxTableWidth}");

            var newTable = new XElement(Namespace.Main + "tbl",
                new XElement(Namespace.Main + "tblPr",
                    new XElement(Namespace.Main + "tblStyle",
                        new XAttribute(Name.MainVal, TableDesign.TableGrid.GetEnumName())),
                    new XElement(Namespace.Main + "tblW", new XAttribute(Namespace.Main + "w", 0),
                        new XAttribute(Namespace.Main + "type", "auto"))));

            WriteTableConditionalFormat(newTable.Element(Namespace.Main + "tblPr"), TableConditionalFormatting.None);

            var grid = new XElement(Namespace.Main + "tblGrid");
            newTable.Add(grid);
            foreach (int width in columnWidths)
            {
                grid.Add(new XElement(Namespace.Main + "gridCol", new XAttribute(Namespace.Main + "w", width * UnitConversion)));
            }

            for (int i = 0; i < rows; i++)
            {
                var row = new XElement(Namespace.Main + "tr");
                foreach (int width in columnWidths)
                    row.Add(CreateTableCell(width));

                newTable.Add(row);
            }
            return newTable;
        }

        /// <summary>
        /// Create and return a cell of a table
        /// </summary>
        public static XElement CreateTableCell(double? width = null)
        {
            string type = width == null ? "auto" : "dxa";
            width ??= 0;

            return new XElement(Namespace.Main + "tc",
                    new XElement(Namespace.Main + "tcPr",
                        new XElement(Namespace.Main + "tcW",
                            new XAttribute(Namespace.Main + "type", type),
                            new XAttribute(Namespace.Main + "w", width.Value * UnitConversion)
                        )),
                    // always has an empty paragraph
                    new XElement(Name.Paragraph,
                        new XAttribute(Name.ParagraphId,
                            HelperFunctions.GenerateHexId()))
            );
        }
    }
}
