using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace DXPlus.Internal;

/// <summary>
/// Helper class to deal with DOCX core and custom properties
/// </summary>
internal abstract class PropertyCollectionHelper<TK,TV> : IDictionary<TK,TV> 
    where TK : notnull
    where TV : class
{
    /// <summary>
    /// Read the package and get all the values.
    /// </summary>
    /// <returns></returns>
    protected abstract Dictionary<TK,TV>? LoadProperties(bool create);

    /// <summary>
    /// Save the dictionary back out to the package.
    /// </summary>
    protected internal abstract void Save();

    /// <summary>
    /// Called before a change is committed.
    /// </summary>
    /// <param name="key"></param>
    /// <param name="value"></param>
    protected virtual void OnKeyValueChanging(TK key, TV? value)
    {
    }

    /// <summary>
    /// Called after property change occurs.
    /// </summary>
    protected abstract void OnKeyValueChanged(TK key, TV? value);

    /// <summary>
    /// Returns an enumerator for the key/value pairs
    /// </summary>
    /// <returns></returns>
    public IEnumerator<KeyValuePair<TK,TV>> GetEnumerator() 
        => (LoadProperties(false) ?? Enumerable.Empty<KeyValuePair<TK,TV>>()).GetEnumerator();

    /// <summary>
    /// Returns an enumerator for the key/value pairs
    /// </summary>
    /// <returns></returns>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>
    /// Add a new value to our document properties.
    /// </summary>
    /// <param name="item"></param>
    /// <exception cref="NotImplementedException"></exception>
    void ICollection<KeyValuePair<TK, TV>>.Add(KeyValuePair<TK, TV> item) => Add(item.Key, item.Value);

    /// <summary>
    /// Clear all document properties.
    /// </summary>
    public void Clear()
    {
        var properties = LoadProperties(false);
        if (properties != null)
        {
            var keys = properties.Keys.ToList();
            foreach (var item in keys)
            {
                OnKeyValueChanging(item, null);
                properties.Remove(item);
                OnKeyValueChanged(item, null);
            }
        }
    }

    /// <summary>
    /// Returns whether the given item is in the dictionary.
    /// </summary>
    /// <param name="item"></param>
    /// <returns></returns>
    bool ICollection<KeyValuePair<TK,TV>>.Contains(KeyValuePair<TK,TV> item) 
        => LoadProperties(false)?.Contains(item) == true;

    /// <summary>
    /// Copy the key value pairs to an array.
    /// </summary>
    /// <param name="array"></param>
    /// <param name="arrayIndex"></param>
    /// <exception cref="ArgumentNullException"></exception>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    void ICollection<KeyValuePair<TK,TV>>.CopyTo(KeyValuePair<TK,TV>[] array, int arrayIndex)
    {
        if (array == null) throw new ArgumentNullException(nameof(array));
        if (arrayIndex < 0) throw new ArgumentOutOfRangeException(nameof(arrayIndex));

        var properties = LoadProperties(false);
        if (properties == null) return;

        if (arrayIndex + properties.Count >= array.Length)
            throw new ArgumentOutOfRangeException(nameof(arrayIndex));

        int i = 0;
        foreach (var item in properties)
        {
            array[i + arrayIndex] = item;
            i++;
        }
    }

    /// <summary>
    /// Remove an item from the dictionary
    /// </summary>
    /// <param name="item"></param>
    /// <returns></returns>
    bool ICollection<KeyValuePair<TK, TV>>.Remove(KeyValuePair<TK, TV> item) => Remove(item.Key);

    /// <summary>
    /// Count of items in the dictionary
    /// </summary>
    public int Count => LoadProperties(false)?.Count ?? 0;

    /// <summary>
    /// True if the dictionary is readonly.
    /// </summary>
    bool ICollection<KeyValuePair<TK,TV>>.IsReadOnly => false;

    /// <summary>
    /// Add a new value to the dictionary
    /// </summary>
    /// <param name="key"></param>
    /// <param name="value"></param>
    public void Add(TK key, TV value)
    {
        if (value == null) throw new ArgumentNullException(nameof(value));

        var properties = LoadProperties(true);

        OnKeyValueChanging(key, value);
        properties!.Add(key,value);
        OnKeyValueChanged(key, value);
    }

    /// <summary>
    /// True if this dictionary contains the given key.
    /// </summary>
    /// <param name="key"></param>
    /// <returns></returns>
    public bool ContainsKey(TK key) => LoadProperties(false)?.ContainsKey(key) == true;

    /// <summary>
    /// Remove a specific item from the dictionary
    /// </summary>
    /// <param name="key"></param>
    /// <returns></returns>
    public bool Remove(TK key)
    {
        var properties = LoadProperties(false);
        if (properties != null)
        {
            OnKeyValueChanging(key, null);
            if (properties.Remove(key))
            {
                OnKeyValueChanged(key, null);
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Return a specific value for a key.
    /// </summary>
    /// <param name="key">Key</param>
    /// <param name="value">Value</param>
    /// <returns></returns>
    public bool TryGetValue(TK key, [MaybeNullWhen(false)] out TV value)
    {
        var properties = LoadProperties(false);
        if (properties != null)
        {
            return properties.TryGetValue(key, out value!);
        }

        value = null;
        return false;
    }

    /// <summary>
    /// Indexer
    /// </summary>
    /// <param name="key"></param>
    /// <returns></returns>
    public TV this[TK key]
    {
        get
        {
            var properties = LoadProperties(false);
            if (properties == null) throw new KeyNotFoundException();
            return properties[key];
        }
        set
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var properties = LoadProperties(true);
            OnKeyValueChanging(key, value);
            properties![key] = value;
            OnKeyValueChanged(key, value);
        }
    }

    /// <summary>
    /// All the keys
    /// </summary>
    public ICollection<TK> Keys => LoadProperties(true)!.Keys;

    /// <summary>
    /// All the values
    /// </summary>
    public ICollection<TV> Values => LoadProperties(true)!.Values;
}