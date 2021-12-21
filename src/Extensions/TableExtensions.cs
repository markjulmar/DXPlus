using System;

namespace DXPlus
{
    /// <summary>
    /// Extensions for the Table (w:tbl) element
    /// </summary>
    public static class TableExtensions
    {
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
        public static Table AutoFit(this Table table, bool value)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            table.AutoFit = value;
            return table;
        }
    }
}
