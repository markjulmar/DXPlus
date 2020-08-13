namespace DXPlus
{
    /// <summary>
    /// Table Cell Border styles
    /// source: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.tablecellborders.aspx
    /// </summary>
    public enum BorderStyle
    {
        Empty   = -1,
        None    = 0,
        Single  = 1,
        Thick,
        DoubleLine,
        Dotted,
        Dashed,
        DotDash,
        DotDotDash,
        Triple,
        ThinThickSmallGap,
        ThickThinSmallGap,
        ThinThickThinSmallGap,
        ThinThickMediumGap,
        ThickThinMediumGap,
        ThinThickThinMediumGap,
        ThinThickLargeGap,
        ThickThinLargeGap,
        ThinThickThinLargeGap,
        Wave,
        DoubleWave,
        DashSmallGap,
        DashDotStroked,
        ThreeDEmboss,
        ThreeDEngrave,
        Outset,
        Inset
    }
}