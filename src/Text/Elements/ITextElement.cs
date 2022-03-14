namespace DXPlus;

/// <summary>
/// Abstraction around child elements of a Document Run.
/// </summary>
public interface ITextElement
{
    /// <summary>
    /// Parent run object
    /// </summary>
    Run? Parent { get; }

    /// <summary>
    /// Name for this element.
    /// </summary>
    string ElementType { get; }
}