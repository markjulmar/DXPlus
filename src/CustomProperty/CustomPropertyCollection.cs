using System.Collections;
using System.Diagnostics;
using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Helper methods to work with custom document properties.
/// </summary>
internal sealed class CustomPropertyCollection : ICustomPropertyCollection
{
    private readonly Package package;
    private readonly Document owner;
    private XDocument? document;
    private XElementCollection<CustomProperty>? propertyCollection;

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="package"></param>
    /// <param name="document"></param>
    /// <exception cref="ArgumentNullException"></exception>
    internal CustomPropertyCollection(Package package, Document document)
    {
        this.package = package ?? throw new ArgumentNullException(nameof(package));
        this.owner = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <summary>
    /// Read the package and get all the values.
    /// </summary>
    /// <returns></returns>
    private IList<CustomProperty>? LoadProperties(bool create)
    {
        if (propertyCollection != null) return propertyCollection;

        PackagePart customPropertiesPart;

        if (!package.PartExists(Relations.CustomProperties.Uri))
        {
            if (!create) return null;

            customPropertiesPart = package.CreatePart(Relations.CustomProperties.Uri, Relations.CustomProperties.ContentType, CompressionOption.Maximum);
            document = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(Namespace.CustomPropertiesSchema + "Properties", 
                    new XAttribute(XNamespace.Xmlns + "vt", Namespace.CustomVTypesSchema))
            );

            customPropertiesPart.Save(document);
            package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, Relations.CustomProperties.RelType);
        }
        else
        {
            customPropertiesPart = package.GetPart(Relations.CustomProperties.Uri);
            document = customPropertiesPart.Load();
        }

        document.Changed += OnDocumentChanged;


        propertyCollection = new XElementCollection<CustomProperty>(document.Root!, null,
            CustomProperty.TagName, xe => new CustomProperty(xe));

        return propertyCollection;
    }

    /// <summary>
    /// Called when the XML document is changed.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void OnDocumentChanged(object? sender, XObjectChangeEventArgs e)
    {
        // A value changed. Update fields.
        foreach (var prop in propertyCollection!)
        {
            owner.UpdateComplexFieldUsage(prop.Name.ToUpper(), prop.Value??"");
        }
    }

    /// <summary>
    /// Save the dictionary back out to the package.
    /// </summary>
    internal void Save()
    {
        // Never loaded?
        if (propertyCollection == null) return;

        Debug.Assert(package.PartExists(Relations.CustomProperties.Uri));
        Debug.Assert(document != null);

        // Get the next property id in the document
        int nextId = Math.Max(2, document.LocalNameDescendants("property")
            .Select(p => int.TryParse(p.AttributeValue("pid"), out var result) ? result : 0)
            .DefaultIfEmpty().Max() + 1);
        
        foreach (var cp in propertyCollection.Where(p => p.Id is null or 0))
            cp.Id = nextId++;

        var customPropertiesPart = package.GetPart(Relations.CustomProperties.Uri);
        customPropertiesPart.Save(document);
    }

    /// <summary>Returns an enumerator that iterates through the collection.</summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<CustomProperty> GetEnumerator() => (LoadProperties(false) ?? Enumerable.Empty<CustomProperty>()).GetEnumerator();

    /// <summary>Returns an enumerator that iterates through a collection.</summary>
    /// <returns>An <see cref="T:System.Collections.IEnumerator" /> object that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Adds an item to the <see cref="T:System.Collections.Generic.ICollection`1" />.</summary>
    /// <param name="item">The object to add to the <see cref="T:System.Collections.Generic.ICollection`1" />.</param>
    /// <exception cref="T:System.NotSupportedException">The <see cref="T:System.Collections.Generic.ICollection`1" /> is read-only.</exception>
    public void Add(CustomProperty item) => LoadProperties(true)!.Add(item);

    /// <summary>Removes all items from the <see cref="T:System.Collections.Generic.ICollection`1" />.</summary>
    /// <exception cref="T:System.NotSupportedException">The <see cref="T:System.Collections.Generic.ICollection`1" /> is read-only.</exception>
    public void Clear() => LoadProperties(false)?.Clear();

    /// <summary>Determines whether the <see cref="T:System.Collections.Generic.ICollection`1" /> contains a specific value.</summary>
    /// <param name="item">The object to locate in the <see cref="T:System.Collections.Generic.ICollection`1" />.</param>
    /// <returns>
    /// <see langword="true" /> if <paramref name="item" /> is found in the <see cref="T:System.Collections.Generic.ICollection`1" />; otherwise, <see langword="false" />.</returns>
    public bool Contains(CustomProperty item) => LoadProperties(false)?.Contains(item) == true;

    /// <summary>Copies the elements of the <see cref="T:System.Collections.Generic.ICollection`1" /> to an <see cref="T:System.Array" />, starting at a particular <see cref="T:System.Array" /> index.</summary>
    /// <param name="array">The one-dimensional <see cref="T:System.Array" /> that is the destination of the elements copied from <see cref="T:System.Collections.Generic.ICollection`1" />. The <see cref="T:System.Array" /> must have zero-based indexing.</param>
    /// <param name="arrayIndex">The zero-based index in <paramref name="array" /> at which copying begins.</param>
    /// <exception cref="T:System.ArgumentNullException">
    /// <paramref name="array" /> is <see langword="null" />.</exception>
    /// <exception cref="T:System.ArgumentOutOfRangeException">
    /// <paramref name="arrayIndex" /> is less than 0.</exception>
    /// <exception cref="T:System.ArgumentException">The number of elements in the source <see cref="T:System.Collections.Generic.ICollection`1" /> is greater than the available space from <paramref name="arrayIndex" /> to the end of the destination <paramref name="array" />.</exception>
    public void CopyTo(CustomProperty[] array, int arrayIndex) => LoadProperties(false)?.CopyTo(array, arrayIndex);

    /// <summary>Removes the first occurrence of a specific object from the <see cref="T:System.Collections.Generic.ICollection`1" />.</summary>
    /// <param name="item">The object to remove from the <see cref="T:System.Collections.Generic.ICollection`1" />.</param>
    /// <exception cref="T:System.NotSupportedException">The <see cref="T:System.Collections.Generic.ICollection`1" /> is read-only.</exception>
    /// <returns>
    /// <see langword="true" /> if <paramref name="item" /> was successfully removed from the <see cref="T:System.Collections.Generic.ICollection`1" />; otherwise, <see langword="false" />. This method also returns <see langword="false" /> if <paramref name="item" /> is not found in the original <see cref="T:System.Collections.Generic.ICollection`1" />.</returns>
    public bool Remove(CustomProperty item) => LoadProperties(false)?.Remove(item) == true;

    /// <summary>Gets the number of elements contained in the <see cref="T:System.Collections.Generic.ICollection`1" />.</summary>
    /// <returns>The number of elements contained in the <see cref="T:System.Collections.Generic.ICollection`1" />.</returns>
    public int Count => LoadProperties(false)?.Count ?? 0;

    /// <summary>Gets a value indicating whether the <see cref="T:System.Collections.Generic.ICollection`1" /> is read-only.</summary>
    /// <returns>
    /// <see langword="true" /> if the <see cref="T:System.Collections.Generic.ICollection`1" /> is read-only; otherwise, <see langword="false" />.</returns>
    public bool IsReadOnly => false;

    /// <summary>Determines the index of a specific item in the <see cref="T:System.Collections.Generic.IList`1" />.</summary>
    /// <param name="item">The object to locate in the <see cref="T:System.Collections.Generic.IList`1" />.</param>
    /// <returns>The index of <paramref name="item" /> if found in the list; otherwise, -1.</returns>
    public int IndexOf(CustomProperty item) => LoadProperties(false)?.IndexOf(item) ?? -1;

    /// <summary>Inserts an item to the <see cref="T:System.Collections.Generic.IList`1" /> at the specified index.</summary>
    /// <param name="index">The zero-based index at which <paramref name="item" /> should be inserted.</param>
    /// <param name="item">The object to insert into the <see cref="T:System.Collections.Generic.IList`1" />.</param>
    /// <exception cref="T:System.ArgumentOutOfRangeException">
    /// <paramref name="index" /> is not a valid index in the <see cref="T:System.Collections.Generic.IList`1" />.</exception>
    /// <exception cref="T:System.NotSupportedException">The <see cref="T:System.Collections.Generic.IList`1" /> is read-only.</exception>
    public void Insert(int index, CustomProperty item) => LoadProperties(true)!.Insert(index, item);

    /// <summary>Removes the <see cref="T:System.Collections.Generic.IList`1" /> item at the specified index.</summary>
    /// <param name="index">The zero-based index of the item to remove.</param>
    /// <exception cref="T:System.ArgumentOutOfRangeException">
    /// <paramref name="index" /> is not a valid index in the <see cref="T:System.Collections.Generic.IList`1" />.</exception>
    /// <exception cref="T:System.NotSupportedException">The <see cref="T:System.Collections.Generic.IList`1" /> is read-only.</exception>
    public void RemoveAt(int index) => (LoadProperties(false) ?? throw new ArgumentOutOfRangeException()).RemoveAt(index);

    /// <summary>Gets or sets the element at the specified index.</summary>
    /// <param name="index">The zero-based index of the element to get or set.</param>
    /// <exception cref="T:System.ArgumentOutOfRangeException">
    /// <paramref name="index" /> is not a valid index in the <see cref="T:System.Collections.Generic.IList`1" />.</exception>
    /// <exception cref="T:System.NotSupportedException">The property is set and the <see cref="T:System.Collections.Generic.IList`1" /> is read-only.</exception>
    /// <returns>The element at the specified index.</returns>
    public CustomProperty this[int index]
    {
        get => (LoadProperties(false) ?? throw new ArgumentOutOfRangeException())[index];
        set => LoadProperties(true)![index] = value;
    }

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    public void Add(string name, string value) => Add(new CustomProperty(name, value));

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    public void Add(string name, double value) => Add(new CustomProperty(name, value));

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    public void Add(string name, bool value) => Add(new CustomProperty(name, value));

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    public void Add(string name, DateTime value) => Add(new CustomProperty(name, value));

    /// <summary>
    /// Add a custom property to this document.
    /// If a custom property already exists with the same name it will be replace.
    /// CustomProperty names are case insensitive.
    /// </summary>
    public void Add(string name, int value) => Add(new CustomProperty(name, value));

    /// <summary>
    /// Locate a custom property in the collection by name
    /// </summary>
    /// <param name="name">Name of the property</param>
    /// <param name="property">Returned property</param>
    /// <returns>True/False if property was located</returns>
    public bool TryGetValue(string name, out CustomProperty? property)
    {
        var item = LoadProperties(false)?.FirstOrDefault(cp => cp.Name == name);
        property = item;
        return item != null;
    }
}