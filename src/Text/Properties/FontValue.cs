using System.Drawing;

namespace DXPlus;

/// <summary>
/// This represents a font family in a Word document.
/// </summary>
public sealed class FontValue : IEquatable<FontValue>
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="name"></param>
    public FontValue(string name)
    {
        Name = name;
    }

    /// <summary>
    /// Family name for this font.
    /// </summary>
    public string Name { get; }

    /// <summary>Returns a string that represents the current object.</summary>
    /// <returns>A string that represents the current object.</returns>
    public override string ToString()
    {
        return Name;
    }

    /// <summary>
    /// Returns all the available font families on this computer.
    /// </summary>
    public static IEnumerable<FontValue> FontFamilies 
        => FontFamily.Families.Select(item => new FontValue(item.Name));

    /// <summary>Indicates whether the current object is equal to another object of the same type.</summary>
    /// <param name="other">An object to compare with this object.</param>
    /// <returns>
    /// <see langword="true" /> if the current object is equal to the <paramref name="other" /> parameter; otherwise, <see langword="false" />.</returns>
    public bool Equals(FontValue? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return Name == other.Name;
    }

    /// <summary>Determines whether the specified object is equal to the current object.</summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>
    /// <see langword="true" /> if the specified object  is equal to the current object; otherwise, <see langword="false" />.</returns>
    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || obj is FontValue other && Equals(other);
    }

    /// <summary>Serves as the default hash function.</summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode()
    {
        return Name.GetHashCode();
    }
}
