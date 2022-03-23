using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Custom file property types expressed in the VariantType namespace.
/// </summary>
public enum CustomPropertyType
{
    /// <summary>
    /// None, empty, null
    /// </summary>
    [XmlAttribute("null")] None = 0,

    /// <summary>
    /// Text values
    /// </summary>
    [XmlAttribute("lpwstr")] Text,

    /// <summary>
    /// COM BSTR values
    /// </summary>
    [XmlAttribute("bstr")] BSTR,

    /// <summary>
    /// System.DateTime
    /// </summary>
    [XmlAttribute("date")] DateTime,

    /// <summary>
    /// Filetime
    /// </summary>
    [XmlAttribute("filetime")] FileTime,

    /// <summary>
    /// System.Byte
    /// </summary>
    [XmlAttribute("u1")] U1,

    /// <summary>
    /// System.UShort
    /// </summary>
    [XmlAttribute("u2")] U2,

    /// <summary>
    /// unsigned 4-byte integer
    /// </summary>
    [XmlAttribute("u4")] U4,

    /// <summary>
    /// Unsigned integer (same storage as U4)
    /// </summary>
    [XmlAttribute("uint")] UnsignedInteger,

    /// <summary>
    /// System.ULong
    /// </summary>
    [XmlAttribute("u8")] U8,

    /// <summary>
    /// System.SByte
    /// </summary>
    [XmlAttribute("i1")] I1,

    /// <summary>
    /// System.Short
    /// </summary>
    [XmlAttribute("i2")] I2,

    /// <summary>
    /// System.Int32
    /// </summary>
    [XmlAttribute("i4")] I4,

    /// <summary>
    /// Integer type (same storage as I4)
    /// </summary>
    [XmlAttribute("int")] Integer,

    /// <summary>
    /// System.Long
    /// </summary>
    [XmlAttribute("i8")] I8,

    /// <summary>
    /// 4-byte real number
    /// </summary>
    [XmlAttribute("r4")] R4,

    /// <summary>
    /// 8-byte real number
    /// </summary>
    [XmlAttribute("r8")] R8,

    /// <summary>
    /// Decimal type
    /// </summary>
    [XmlAttribute("decimal")] Decimal,

    /// <summary>
    ///  0xHHHHHHHH error code
    /// </summary>
    [XmlAttribute("error")] ErrorCode,

    /// <summary>
    /// System.GUID
    /// </summary>
    [XmlAttribute("clsid")] CLSID,

    /// <summary>
    /// System.Boolean
    /// </summary>
    [XmlAttribute("bool")] Boolean,

    /// <summary>
    /// Decimal with 4 digits (\s*[0-9]*\.[0-9]{4}\s*)
    /// </summary>
    [XmlAttribute("cy")] Currency,

}