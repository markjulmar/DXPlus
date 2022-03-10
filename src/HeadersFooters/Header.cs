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
}