using System.Runtime.Serialization;

namespace DXPlus;

/// <summary>
/// Exception used when an unexpected token or missing token is found in the document structure.
/// </summary>
[Serializable]
public class DocumentFormatException : Exception
{
    /// <summary>
    /// Name of the element
    /// </summary>
    public string? Element { get; init; }

    /// <summary>
    /// Constructor
    /// </summary>
    public DocumentFormatException()
    {
    }

    /// <summary>
    /// constructor
    /// </summary>
    /// <param name="elementName">Document element name</param>
    public DocumentFormatException(string elementName)
    {
        Element = elementName;
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="elementName"></param>
    /// <param name="message"></param>
    public DocumentFormatException(string elementName, string message) : base(message)
    {
        Element = elementName;
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="elementName"></param>
    /// <param name="message"></param>
    /// <param name="inner"></param>
    public DocumentFormatException(string elementName, string message, Exception inner) : base(message, inner)
    {
        Element = elementName;
    }
}