using System;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    /// <summary>
    /// Helper methods to create/manage table (w:tbl) elements
    /// </summary>
    internal static class TableHelpers
    {
        private const int MaxTableWidth = 8192;
        public const double UnitConversion = 20.0;  // Height/Widths are represented in 20ths of a pt.

        public static XElement CreateTable(int rows, int columns)
        {
            if (rows < 1)
                throw new ArgumentOutOfRangeException(nameof(rows), "Cannot be < 1");
            if (columns < 1)
                throw new ArgumentOutOfRangeException(nameof(columns), "Cannot be < 1");

            int[] columnWidths = new int[columns];
            Array.Fill(columnWidths, MaxTableWidth / columns);
            return CreateTable(rows, columnWidths);
        }

        public static XElement CreateTable(int rows, int[] columnWidths)
        {
            var newTable = new XElement(Namespace.Main + "tbl",
                new XElement(Namespace.Main + "tblPr",
                        new XElement(Namespace.Main + "tblStyle", new XAttribute(Name.MainVal, TableDesign.TableGrid.GetEnumName())),
                        new XElement(Namespace.Main + "tblW", new XAttribute(Namespace.Main + "w", MaxTableWidth),
                                new XAttribute(Namespace.Main + "type", "dxa")),
                        new XElement(Namespace.Main + "tblLook", new XAttribute(Name.MainVal, "04A0"))
                )
            );

            for (int i = 0; i < rows; i++)
            {
                var row = new XElement(Namespace.Main + "tr");
                foreach (var width in columnWidths)
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
