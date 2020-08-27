using System;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
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
        internal static XElement CreateTableCell(double? width = null)
        {
            string type = width == null ? "auto" : "dxa";
            width ??= 0;

            return new XElement(Namespace.Main + "tc",
                    new XElement(Namespace.Main + "tcPr",
                        new XElement(Namespace.Main + "tcW",
                            new XAttribute(Namespace.Main + "type", type),
                            new XAttribute(Namespace.Main + "w", width.Value * UnitConversion)
                        )),
                    new XElement(Name.Paragraph) // always has an empty paragraph
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


        /// <summary>
        /// Add a new table after this container
        /// </summary>
        /// <param name="container">Container owner</param>
        /// <param name="rows">Number of rows</param>
        /// <param name="columns">Number of columns</param>
        public static Table AddTableAfterSelf(this InsertBeforeOrAfter container, int rows, int columns) => container.AddTableAfterSelf(new Table(rows, columns));

        /// <summary>
        /// Insert a new table before this container
        /// </summary>
        /// <param name="container">Container owner</param>
        /// <param name="rows">Number of rows</param>
        /// <param name="columns">Number of columns</param>
        public static Table InsertTableBeforeSelf(this InsertBeforeOrAfter container, int rows, int columns) => container.InsertTableBeforeSelf(new Table(rows, columns));
    }
}
