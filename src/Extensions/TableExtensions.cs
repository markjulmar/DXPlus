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
        if (table == null) throw new ArgumentNullException(nameof(table));
        table.Properties.Alignment = value;
        return table;
    }

    /// <summary>
    /// Fluent syntax for AutoFit
    /// </summary>
    /// <param name="table"></param>
    /// <returns></returns>
    public static Table AutoFit(this Table table)
    {
        if (table == null) throw new ArgumentNullException(nameof(table));
        table.Properties.TableLayout = TableLayout.AutoFit;
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
        if (table == null) throw new ArgumentNullException(nameof(table));
        table.Properties.SetOutsideBorders(border);
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
        if (table == null) throw new ArgumentNullException(nameof(table));
        table.Properties.SetInsideBorders(border);
        return table;
    }

    /// <summary>
    /// Set all the outside borders of the table
    /// </summary>
    /// <param name="properties"></param>
    /// <param name="border"></param>
    /// <returns></returns>
    public static void SetOutsideBorders(this TableProperties properties, Border border)
    {
        if (properties == null) throw new ArgumentNullException(nameof(properties));

        properties.LeftBorder = border;
        properties.RightBorder = border;
        properties.TopBorder = border;
        properties.BottomBorder = border;
    }

    /// <summary>
    /// Set all the outside borders of the table
    /// </summary>
    /// <param name="properties"></param>
    /// <param name="border"></param>
    /// <returns></returns>
    public static void SetInsideBorders(this TableProperties properties, Border border)
    {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        properties.InsideHorizontalBorder = border;
        properties.InsideVerticalBorder = border;
    }

}