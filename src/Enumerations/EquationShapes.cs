using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Shapes usable in an equation
/// </summary>
public enum EquationShapes
{
    /// <summary>
    /// Plus sign
    /// </summary>
    [XmlAttribute("mathPlus")] Plus,

    /// <summary>
    /// Minus sign
    /// </summary>
    [XmlAttribute("mathMinus")] Minus,

    /// <summary>
    /// Multiplication
    /// </summary>
    [XmlAttribute("mathMultiply")] Multiply,

    /// <summary>
    /// Division
    /// </summary>
    [XmlAttribute("mathDivide")] Divide,

    /// <summary>
    /// Equality
    /// </summary>
    [XmlAttribute("mathEqual")] Equal,

    /// <summary>
    /// Not equal
    /// </summary>
    [XmlAttribute("mathNotEqual")] NotEqual
};