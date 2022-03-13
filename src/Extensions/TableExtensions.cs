namespace DXPlus;

/// <summary>
/// ImageExtensions for the Table (w:tbl) element
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
    /// <returns></returns>
    public static Table AutoFit(this Table table)
    {
        if (table == null)
            throw new ArgumentNullException(nameof(table));

        table.TableLayout = TableLayout.AutoFit;
        return table;
    }

    /// <summary>
    /// Set all the outside borders of the table
    /// </summary>
    /// <param name="table"></param>
    /// <param name="border"></param>
    /// <returns></returns>
    public static Table SetOutsideBorders(this Table table, Border border)
    {
        table.LeftBorder = border;
        table.RightBorder = border;
        table.TopBorder = border;
        table.BottomBorder = border;
        return table;
    }

    /// <summary>
    /// Set all the outside borders of the table
    /// </summary>
    /// <param name="table"></param>
    /// <param name="border"></param>
    /// <returns></returns>
    public static Table SetInsideBorders(this Table table, Border border)
    {
        table.InsideHorizontalBorder = border;
        table.InsideVerticalBorder = border;
        return table;
    }

}