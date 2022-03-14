namespace DXPlus;

/// <summary>
/// This ties a comment to a range set.
/// </summary>
public sealed class CommentRange
{
    /// <summary>
    /// The comment
    /// </summary>
    public Comment Comment { get; }

    /// <summary>
    /// The paragraph owner
    /// </summary>
    public Paragraph Owner { get; }

    /// <summary>
    /// Starting run
    /// </summary>
    public Run? RangeStart { get; }

    /// <summary>
    /// Ending run - might be the same as start
    /// </summary>
    public Run? RangeEnd { get; }

    /// <summary>
    /// Constructor
    /// </summary>
    internal CommentRange(Paragraph owner, Run? start, Run? end, Comment comment)
    {
        this.Owner = owner;
        this.RangeStart = start;
        this.RangeEnd = end;
        this.Comment = comment;
    }
}