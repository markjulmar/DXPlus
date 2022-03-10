namespace DXPlus;

/// <summary>
/// How to match formatting when doing find/replace in text.
/// </summary>
public enum MatchFormattingOptions
{
    /// <summary>
    /// Require exact match
    /// </summary>
    ExactMatch,
    
    /// <summary>
    /// Allow subset of formatting to match
    /// </summary>
    SubsetMatch
};