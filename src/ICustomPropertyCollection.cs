namespace DXPlus;

/// <summary>
/// Interface for custom property helpers.
/// </summary>
public interface ICustomPropertyCollection : IList<CustomProperty>
{
    /// <summary>
    /// Locate a custom property in the collection by name
    /// </summary>
    /// <param name="name">Name of the property</param>
    /// <param name="property">Returned property</param>
    /// <returns>True/False if property was located</returns>
    bool TryGetValue(string name, out CustomProperty? property);

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    void Add(string name, string value);

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    void Add(string name, double value);

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    void Add(string name, bool value);

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    void Add(string name, DateTime value);

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    void Add(string name, int value);
}