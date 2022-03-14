namespace DXPlus;

/// <summary>
/// Represents a single header (even, odd, first) in a document
/// </summary>
public sealed class Header : HeaderOrFooter, IEquatable<Header>
{
    /// <summary>
    /// Determines equality for a header
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Header? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a header
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Header);

    /// <summary>
    /// Returns hashcode for this header
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();
}