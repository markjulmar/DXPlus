using System.Collections;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Base collection class to manage a collection of children.
/// </summary>
/// <typeparam name="T">Child type</typeparam>
public class XElementCollection<T> : IList<T>, IReadOnlyList<T> where T : XElementWrapper
{
    private readonly XName? parentTag;
    private readonly XName childTag;
    private readonly Func<XElement, T> createChildFunc;
    private readonly bool isReadOnly;
    private readonly XElement parent;

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="parent"></param>
    /// <param name="parentTag"></param>
    /// <param name="childTag"></param>
    /// <param name="createChildFunc"></param>
    /// <param name="isReadOnly"></param>
    internal XElementCollection(XElement parent, XName? parentTag, XName childTag, Func<XElement, T> createChildFunc, bool isReadOnly = false)
    {
        this.parent = parent;
        this.parentTag = parentTag;
        this.childTag = childTag;
        this.createChildFunc = createChildFunc;
        this.isReadOnly = isReadOnly;
    }

    /// <inheritdoc />
    public IEnumerator<T> GetEnumerator()
    {
        var root = parentTag != null ? parent.Element(parentTag) : parent;
        return (root?.Elements(childTag).Select(el => createChildFunc(el))
                ?? Enumerable.Empty<T>()).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    /// <summary>
    /// Returns the array element in the XML structure.
    /// </summary>
    /// <param name="create">True to create if it doesn't exist.</param>
    /// <returns>Element</returns>
    private XElement? ArrayElement(bool create)
    {
        if (parentTag != null)
        {
            return create ? parent.GetOrAddElement(parentTag) : parent.Element(parentTag);
        }

        return parent;
    }

    /// <inheritdoc />
    public void Add(T item)
    {
        if (isReadOnly) throw new NotSupportedException("Collection is read-only.");
        if (item == null) throw new ArgumentNullException(nameof(item));

        var xe = item.Xml;
        if (xe == null) throw new ArgumentException($"{item.GetType().Name} is invalid.", nameof(item));
        if (xe.Parent != null) throw new ArgumentException($"{item.GetType().Name} is already in document.", nameof(item));
        if (xe.Name != childTag) throw new ArgumentException($"{item.GetType().Name} does not have the expected XML tag {childTag}");

        ArrayElement(true)!.Add(xe);
    }

    /// <inheritdoc />
    public void Clear()
    {
        if (isReadOnly) throw new NotSupportedException("Collection is read-only.");
        if (parentTag != null)
        {
            // Remove the parent and all children.
            parent.Element(parentTag)?.Remove();
        }
        else
        {
            // Just remove children
            parent.Elements().Remove();
        }
    }

    private void CheckIfDeleteParent()
    {
        if (parentTag != null)
        {
            var root = ArrayElement(false);
            if (root is {HasElements: false})
                root.Remove();
        }
    }

    /// <inheritdoc />
    public bool Contains(T item) => this.IndexOf(item) >= 0;

    /// <inheritdoc />
    public void CopyTo(T[] array, int arrayIndex)
    {
        int pos = 0;
        using var enumerator = GetEnumerator();
        while (enumerator.MoveNext())
        {
            array[arrayIndex + pos] = enumerator.Current;
            pos++;
        }
    }

    /// <inheritdoc />
    public bool Remove(T item)
    {
        if (isReadOnly) throw new NotSupportedException("Collection is read-only.");
        if (item == null) throw new ArgumentNullException(nameof(item));

        int pos = IndexOf(item);
        if (pos >= 0)
        {
            RemoveAt(pos);
            return true;
        }

        return false;
    }

    /// <inheritdoc />
    public int Count => ArrayElement(false)?.Elements(childTag).Count() ?? 0;

    bool ICollection<T>.IsReadOnly => isReadOnly;

    /// <inheritdoc />
    public int IndexOf(T item)
    {
        if (item == null) throw new ArgumentNullException(nameof(item));
        var array = ArrayElement(false);

        var iq = item as IEquatable<T>;

        if (array != null)
        {
            var items = this.ToArray();
            for (int i = 0; i < items.Length; i++)
            {
                var check = items[i];
                if (iq?.Equals(check) == true 
                    || check == item) return i;
            }
        }
        return -1;
    }

    /// <inheritdoc />
    public void Insert(int index, T item)
    {
        if (isReadOnly) throw new NotSupportedException("Collection is read-only.");
        if (item == null) throw new ArgumentNullException(nameof(item));
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));

        var xe = item.Xml;
        if (xe == null) throw new ArgumentException($"{item.GetType().Name} is invalid.", nameof(item));
        if (xe.Parent != null) throw new ArgumentException($"{item.GetType().Name} is already in document.", nameof(item));

        var array = ArrayElement(true);
        var children = array!.Elements(childTag).ToList();
        if (children.Count <= index) throw new ArgumentOutOfRangeException(nameof(index));

        if (index == 0)
        {
            array.AddFirst(xe);
        }
        else
        {
            var before = children[index - 1];
            before.AddAfterSelf(xe);
        }
    }

    /// <inheritdoc />
    public void RemoveAt(int index)
    {
        if (isReadOnly) throw new NotSupportedException("Collection is read-only.");
        if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
        var array = ArrayElement(true);
        var children = array!.Elements(childTag).ToList();
        if (children.Count < index) throw new ArgumentOutOfRangeException(nameof(index));

        children[index].Remove();
        CheckIfDeleteParent();
    }

    /// <inheritdoc />
    public T this[int index]
    {
        get
        {
            using var enumerator = GetEnumerator();
            while (enumerator.MoveNext() && index >= 0)
            {
                if (index == 0) 
                    return enumerator.Current;
                index--;
            }

            throw new ArgumentOutOfRangeException(nameof(index));
        }
        set
        {
            if (isReadOnly) throw new NotSupportedException("Collection is read-only.");
            if (index < 0) throw new ArgumentOutOfRangeException(nameof(index));
            if (value == null) throw new ArgumentNullException(nameof(value));

            var xe = value.Xml;
            if (xe == null) throw new ArgumentException($"{value.GetType().Name} is invalid.", nameof(value));
            if (xe.Parent != null) throw new ArgumentException($"{value.GetType().Name} is already in document.", nameof(value));

            var array = ArrayElement(true);
            var children = array!.Elements(childTag).ToList();
            if (children.Count < index) throw new ArgumentOutOfRangeException(nameof(index));

            children[index].AddAfterSelf(xe);
            children[index].Remove();
        }
    }
}