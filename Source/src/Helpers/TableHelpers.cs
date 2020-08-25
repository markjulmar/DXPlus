using System;
using System.Xml.Linq;

namespace DXPlus.Helpers
{
    public static class TableHelpers
    {
        const int MaxTableWidth = 8192;
        internal const double UnitConversion = 20.0;  // Height/Widths are represented in 20ths of a pt.

        internal static XElement CreateTable(int rows, int columns)
        {
            if (rows < 1)
                throw new ArgumentOutOfRangeException(nameof(rows), "Cannot be < 1");
            if (columns < 1)
                throw new ArgumentOutOfRangeException(nameof(columns), "Cannot be < 1");

            int[] columnWidths = new int[columns];
            Array.Fill(columnWidths, MaxTableWidth / columns);
            return CreateTable(rows, columnWidths);
        }

        internal static XElement CreateTable(int rows, int[] columnWidths)
        {
            var newTable = new XElement(DocxNamespace.Main + "tbl",
                new XElement(DocxNamespace.Main + "tblPr",
                        new XElement(DocxNamespace.Main + "tblStyle", new XAttribute(DocxNamespace.Main + "val", TableDesign.TableGrid.GetEnumName())),
                        new XElement(DocxNamespace.Main + "tblW", new XAttribute(DocxNamespace.Main + "w", MaxTableWidth),
                                new XAttribute(DocxNamespace.Main + "type", "dxa")),
                        new XElement(DocxNamespace.Main + "tblLook", new XAttribute(DocxNamespace.Main + "val", "04A0"))
                )
            );

            for (int i = 0; i < rows; i++)
            {
                var row = new XElement(DocxNamespace.Main + "tr");
                foreach (var width in columnWidths)
                    row.Add(CreateTableCell(width));

                newTable.Add(row);
            }
            return newTable;
        }

        /// <summary>
        /// Create and return a cell of a table
        /// </summary>
        internal static XElement CreateTableCell(double? width = null)
        {
            string type = width == null ? "auto" : "dxa";
            width ??= 0;

            return new XElement(DocxNamespace.Main + "tc",
                    new XElement(DocxNamespace.Main + "tcPr",
                        new XElement(DocxNamespace.Main + "tcW",
                            new XAttribute(DocxNamespace.Main + "type", type),
                            new XAttribute(DocxNamespace.Main + "w", width.Value * UnitConversion)
                        )),
                    new XElement(DocxNamespace.Main + "p") // always has an empty paragraph
            );
        }

        /// <summary>
        /// Fluent syntax for alignment
        /// </summary>
        /// <param name="table"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Table Alignment(this Table table, Alignment value)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));
            
            table.Alignment = value;
            return table;
        }

        /// <summary>
        /// Fluent syntax for AutoFit
        /// </summary>
        /// <param name="table"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Table AutoFit(this Table table, AutoFit value)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            table.AutoFit = value;
            return table;
        }

    }
}
