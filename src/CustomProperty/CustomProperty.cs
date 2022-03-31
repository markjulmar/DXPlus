using System.Globalization;
using System.Xml.Linq;
using DXPlus.Internal;

namespace DXPlus;

/// <summary>
/// Represents a single custom property defined in the document.
/// </summary>
public sealed class CustomProperty : XElementWrapper
{
    internal static readonly XName TagName = Namespace.CustomPropertiesSchema + "property";

    /// <summary>
    /// Override nullable base.
    /// </summary>
    private new XElement Xml => base.Xml!;

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="element"></param>
    internal CustomProperty(XElement element)
    {
        base.Xml = element ?? throw new ArgumentNullException(nameof(element));
        if (Xml.Name != TagName)
            throw new ArgumentException($"Property XML node name ({element.Name}) != {Namespace.CustomPropertiesSchema + "property"}", nameof(element));
    }

    /// <summary>
    /// Public constructor for text properties
    /// </summary>
    public CustomProperty(string name, string value) : this(name, CustomPropertyType.Text, value)
    {
    }

    /// <summary>
    /// Public constructor for boolean properties
    /// </summary>
    public CustomProperty(string name, bool value) : this(name, CustomPropertyType.Boolean, value)
    {
    }

    /// <summary>
    /// Public constructor for GUID properties
    /// </summary>
    public CustomProperty(string name, Guid value) : this(name, CustomPropertyType.CLSID, value)
    {
    }

    /// <summary>
    /// Public constructor for DateTime properties
    /// </summary>
    public CustomProperty(string name, DateTime value) : this(name, CustomPropertyType.DateTime, value)
    {
    }

    /// <summary>
    /// Public constructor for real numbers
    /// </summary>
    public CustomProperty(string name, double value) : this(name, CustomPropertyType.R8, value)
    {
    }

    /// <summary>
    /// Public constructor for real numbers
    /// </summary>
    public CustomProperty(string name, decimal value) : this(name, CustomPropertyType.Decimal, value)
    {
    }

    /// <summary>
    /// Public constructor for integers numbers
    /// </summary>
    public CustomProperty(string name, int value) : this(name, CustomPropertyType.Integer, value)
    {
    }

    /// <summary>
    /// Public constructor for unsigned integers numbers
    /// </summary>
    public CustomProperty(string name, uint value) : this(name, CustomPropertyType.UnsignedInteger, value)
    {
    }

    /// <summary>
    /// Common constructor when creating a new property
    /// </summary>
    /// <param name="name">Name</param>
    /// <param name="propertyType">Type</param>
    /// <param name="value">Value</param>
    public CustomProperty(string name, CustomPropertyType propertyType, object value)
    {
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (value == null) throw new ArgumentNullException(nameof(value));

        string type = propertyType.GetEnumName();
        string tv = CastBasedOnType(propertyType, value) ?? "";
        base.Xml = new XElement(TagName,
            new XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"),
            new XAttribute("name", name),
            new XElement(Namespace.CustomVTypesSchema + type, tv));
    }

    /// <summary>
    /// Unique identifier for the custom property.
    /// </summary>
    public int? Id
    {
        get => int.TryParse(Xml.AttributeValue("pid"), out var result) ? result : null;
        internal set
        {
            if (value < 2) throw new ArgumentOutOfRangeException(nameof(value));
            Xml.SetAttributeValue("pid", value == null ? "" : value.ToString());
        }
    }

    /// <summary>
    /// Name of this custom property
    /// </summary>
    public string Name => Xml.AttributeValue("name")!;

    /// <summary>
    /// Name of a bookmark in the current document from which the value of this custom document property should be extracted.
    /// If this value is present, then any data value should be considered a cache and replaced with the value of this
    /// bookmark (if present) on save. If the bookmark is not present, then this link shall be considered broken
    /// and the cached value shall be retained.
    /// </summary>
    public string? LinkTarget => Xml.AttributeValue("linkTarget");

    /// <summary>
    /// Type of value this custom property can return
    /// </summary>
    public CustomPropertyType Type => IdentifyPropertyType();

    /// <summary>
    /// Return the value as a string
    /// </summary>
    public string? Value
    {
        get => DataNode?.Value;
        set => DataNode!.Value = value??"";
    }

    /// <summary>
    /// Return the value as a specific data type.
    /// </summary>
    public T? As<T>() where T : struct
    {
        string? text = Value;
        if (text == null) return null;
        return GenericHelper<T>.TryParse?.Invoke(text, out var result) == true ? result : null;
    }

    /// <summary>
    /// Override specific for string (ref type).
    /// </summary>
    /// <returns></returns>
    public string? As() => Value;

    /// <summary>
    /// Set the value with error checking.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="value"></param>
    public void SetValue<T>(T value) where T : struct
    {
        Value = CastBasedOnType(Type, value);
    }

    /// <summary>
    /// Set the value with error checking.
    /// </summary>
    /// <param name="type"></param>
    /// <param name="value"></param>
    private static string? CastBasedOnType(CustomPropertyType type, object value)
    {
        switch (type)
        {
            case CustomPropertyType.None: return string.Empty;
            case CustomPropertyType.Text: return value.ToString();
            case CustomPropertyType.BSTR: return value.ToString();
            case CustomPropertyType.DateTime:
            case CustomPropertyType.FileTime:
                if (value is DateTime dt)
                    return dt.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
                break;
            case CustomPropertyType.U1: return Convert.ToByte(value).ToString();
            case CustomPropertyType.U2: return Convert.ToUInt16(value).ToString();
            case CustomPropertyType.U4: return Convert.ToUInt32(value).ToString();
            case CustomPropertyType.UnsignedInteger: return Convert.ToUInt32(value).ToString();
            case CustomPropertyType.U8: return Convert.ToUInt64(value).ToString();
            case CustomPropertyType.I1: return Convert.ToSByte(value).ToString();
            case CustomPropertyType.I2: return Convert.ToInt16(value).ToString();
            case CustomPropertyType.I4: return Convert.ToInt32(value).ToString();
            case CustomPropertyType.Integer: return Convert.ToInt32(value).ToString();
            case CustomPropertyType.I8: return Convert.ToInt64(value).ToString();
            case CustomPropertyType.R4: return Convert.ToSingle(value).ToString(CultureInfo.InvariantCulture);
            case CustomPropertyType.R8: return Convert.ToDouble(value).ToString(CultureInfo.InvariantCulture);
            case CustomPropertyType.Decimal: return Convert.ToDecimal(value).ToString(CultureInfo.InvariantCulture);
            case CustomPropertyType.ErrorCode:
            {
                if (value is string s)
                    value = Convert.ToUInt32(s, 16);
                return $"0x{Convert.ToUInt32(value):X}";
            }
            case CustomPropertyType.Currency: return Convert.ToDecimal(value).ToString("F");
            case CustomPropertyType.CLSID:
                if (value is Guid g)
                    return g.ToString("B");
                break;
            case CustomPropertyType.Boolean:
                return value.ToString() == "1" ||
                        string.Compare(value.ToString(), "true", StringComparison.InvariantCultureIgnoreCase) == 0
                    ? "true"
                    : "false";
        }

        throw new ArgumentOutOfRangeException(nameof(value), $"Expected {type}");
    }

    /// <summary>
    /// Override for SetValue to use a string (ref type).
    /// </summary>
    /// <param name="value"></param>
    public void SetValue(string value) => Value = value;

    /// <summary>
    /// Helper class to generate delegate callbacks using reflection.
    /// Caches off in static field so we only incur the reflection cost once per run.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    static class GenericHelper<T>
    {
        public delegate bool TryParseFunc(string str, out T result);
        private static TryParseFunc? _tryParse;
        public static TryParseFunc? TryParse =>
            _tryParse ??= Delegate.CreateDelegate(typeof(TryParseFunc), typeof(T), "TryParse") as TryParseFunc;
    }

    /// <summary>
    /// Returns the data node for this custom property.
    /// </summary>
    private XElement? DataNode => Xml.Elements().SingleOrDefault(e => e.Name.Namespace == Namespace.CustomVTypesSchema);

    /// <summary>
    /// Locate the data node for this Custom property and parse the type.
    /// </summary>
    /// <returns></returns>
    private CustomPropertyType IdentifyPropertyType()
    {
        var dataNode = DataNode;
        if (dataNode != null)
        {
            string dnName = dataNode.Name.LocalName;
            return Enum.GetValues<CustomPropertyType>()
                .FirstOrDefault(pt => pt.GetEnumName() == dnName);
        }
        return CustomPropertyType.None;
    }

    /// <summary>
    /// Override for ToString
    /// </summary>
    /// <returns></returns>
    public override string ToString() => $"{Name}={Value}";
}