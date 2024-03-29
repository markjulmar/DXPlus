﻿namespace DXPlus;

/// <summary>
/// Represents a footer in a section of the document.
/// This can be default, even, or first.
/// </summary>
public sealed class Footer : HeaderOrFooter, IEquatable<Footer>
{
    /// <summary>
    /// Determines equality for a footer
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(Footer? other) 
        => other != null && (ReferenceEquals(this, other) || Xml == other.Xml);

    /// <summary>
    /// Determines equality for a footer
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public override bool Equals(object? other) => Equals(other as Footer);

    /// <summary>
    /// Returns hashcode for this footer
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode() => Xml.GetHashCode();
}