using System.Diagnostics;
using System.IO.Packaging;
using System.Xml.Linq;
using DXPlus.Internal;
using DXPlus.Resources;

namespace DXPlus;

/// <summary>
/// Interface we expose for core property collection.
/// </summary>
public sealed class CoreProperties
{
    private readonly Package package;
    private readonly Document owner;
    private readonly XDocument document; 

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="package"></param>
    /// <param name="document"></param>
    /// <exception cref="ArgumentNullException"></exception>
    internal CoreProperties(Package package, Document document)
    {
        this.package = package ?? throw new ArgumentNullException(nameof(package));
        this.owner = document ?? throw new ArgumentNullException(nameof(document));
        this.document = LoadCorePropertyDocument();
    }

    /// <summary>
    /// Load or create the core.xml document from the Word package.
    /// </summary>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    private XDocument LoadCorePropertyDocument()
    {
        XDocument corePropDoc;
        if (!package.PartExists(Relations.CoreProperties.Uri))
        {
            _ = CreateCoreProperties(package, out corePropDoc);
        }
        else
        {
            corePropDoc = package.GetPart(Relations.CoreProperties.Uri).Load();
            if (corePropDoc.Root == null || corePropDoc.Root.Name.LocalName != "coreProperties")
                throw new InvalidOperationException("Core.xml malformed.");
        }

        return corePropDoc;
    }

    /// <summary>
    /// Save the core.xml document back to the Word doc package.
    /// </summary>
    internal void Save()
    {
        var corePropPart = package.GetPart(Relations.CoreProperties.Uri);
        if (corePropPart == null) throw new InvalidOperationException("Unable to locate core.xml package part.");
        corePropPart.Save(document);
    }

    /// <summary>
    /// Create the XML document for core document properties.
    /// </summary>
    /// <param name="package"></param>
    /// <param name="corePropDoc"></param>
    /// <returns></returns>
    internal static PackagePart CreateCoreProperties(Package package, out XDocument corePropDoc)
    {
        string userName = Environment.UserInteractive ? Environment.UserName : string.Empty;
        if (string.IsNullOrWhiteSpace(userName)) userName = "Office User";

        var corePropPart = package.CreatePart(Relations.CoreProperties.Uri, Relations.CoreProperties.ContentType, CompressionOption.Maximum);
        corePropDoc = Resource.CorePropsXml(userName, DateTime.UtcNow);
        Debug.Assert(corePropDoc.Root != null);

        corePropPart.Save(corePropDoc);
        package.CreateRelationship(corePropPart.Uri, TargetMode.Internal, Relations.CoreProperties.RelType);
        return corePropPart;
    }

    /// <summary>
    /// Get a value from the document
    /// </summary>
    /// <param name="propertyName"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    internal string? GetValue(DocumentPropertyName propertyName)
    {
        Debug.Assert(document != null);
        string key = propertyName.GetEnumName();

        if (!SplitXmlName(key, out var ns, out var localName))
            ns = "cp"; // default

        var xns = document.Root!.GetNamespaceOfPrefix(ns);
        if (xns == null)
            throw new InvalidOperationException($"Unable to identify namespace {ns} used core property {localName}.");
        
        return document.Root!.Element(xns + localName)?.Value;
    }

    /// <summary>
    /// Set a value into the document
    /// </summary>
    /// <param name="propertyName"></param>
    /// <param name="value"></param>
    /// <exception cref="InvalidOperationException"></exception>
    private void SetValue(DocumentPropertyName propertyName, string value)
    {
        if (value == null) throw new ArgumentNullException(nameof(value));
        string key = propertyName.GetEnumName();

        if (!SplitXmlName(key, out var ns, out var localName))
            ns = "cp"; // default

        var xns = document.Root!.GetNamespaceOfPrefix(ns);
        if (xns == null)
            throw new InvalidOperationException($"Unable to identify namespace {ns} used core property {localName}.");

        var element = document.Root.Element(xns + localName);
        if (element != null)
        {
            element.Value = value;
        }
        else
        {
            var ne = new XElement(xns + localName, value);
            if (propertyName is DocumentPropertyName.CreatedDate
                or DocumentPropertyName.SaveDate)
            {
                // Special case for dates
                XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
                ne.Add(new XAttribute(xsi+"type", "dcterms:W3CDTF"));
            }

            document.Root.Add(ne);
        }

        // Update any field usage
        owner.UpdateComplexFieldUsage(propertyName.ToString().ToUpper(), value);
    }

    /// <summary>
    /// This helper splits an XML name into a namespace + localName.
    /// </summary>
    /// <param name="name">Full name</param>
    /// <param name="ns">Returning namespace</param>
    /// <param name="localName">Returning localName</param>
    private static bool SplitXmlName(string name, out string ns, out string localName)
    {
        if (name.Contains(':'))
        {
            var parts = name.Split(':');
            ns = parts[0];
            localName = parts[1];
            return true;
        }

        ns = string.Empty;
        localName = name;
        return false;
    }

    /// <summary>
    /// Title
    /// </summary>
    public string? Title
    {
        get => GetValue(DocumentPropertyName.Title);
        set => SetValue(DocumentPropertyName.Title, value ?? string.Empty);
    }

    /// <summary>
    /// Subject
    /// </summary>
    public string? Subject
    {
        get => GetValue(DocumentPropertyName.Subject);
        set => SetValue(DocumentPropertyName.Subject, value ?? string.Empty);
    }

    /// <summary>
    /// The creator
    /// </summary>
    public string? Creator
    {
        get => GetValue(DocumentPropertyName.Creator);
        set => SetValue(DocumentPropertyName.Creator, value ?? string.Empty);
    }

    /// <summary>
    /// Keywords
    /// </summary>
    public string? Keywords
    {
        get => GetValue(DocumentPropertyName.Keywords);
        set => SetValue(DocumentPropertyName.Keywords, value ?? string.Empty);
    }

    /// <summary>
    /// Description/Comments
    /// </summary>
    public string? Description
    {
        get => GetValue(DocumentPropertyName.Description);
        set => SetValue(DocumentPropertyName.Description, value ?? string.Empty);
    }

    /// <summary>
    /// Last modified by author
    /// </summary>
    public string? LastSavedBy
    {
        get => GetValue(DocumentPropertyName.LastSavedBy);
        set => SetValue(DocumentPropertyName.LastSavedBy, value ?? string.Empty);
    }

    /// <summary>
    /// Revision/Version
    /// </summary>
    public string? Revision
    {
        get => GetValue(DocumentPropertyName.Revision);
        set => SetValue(DocumentPropertyName.Revision, value ?? string.Empty);
    }

    /// <summary>
    /// Category
    /// </summary>
    public string? Category
    {
        get => GetValue(DocumentPropertyName.Category);
        set => SetValue(DocumentPropertyName.Category, value ?? string.Empty);
    }

    /// <summary>
    /// Date owner was created
    /// </summary>
    public DateTime? CreatedDate
    {
        get
        {
            var value = GetValue(DocumentPropertyName.CreatedDate);
            return DateTime.TryParse(value, out var dt) ? dt.ToLocalTime() : null;
        }
        set
        {
            var dt = (value ?? DateTime.Now).ToUniversalTime();
            SetValue(DocumentPropertyName.CreatedDate, dt.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }
    }

    /// <summary>
    /// Last date/time owner was saved
    /// </summary>
    public DateTime? SaveDate
    {
        get
        {
            var value = GetValue(DocumentPropertyName.SaveDate);
            return DateTime.TryParse(value, out var dt) ? dt.ToLocalTime() : null;
        }
        set
        {
            var dt = (value ?? DateTime.Now).ToUniversalTime();
            SetValue(DocumentPropertyName.SaveDate, dt.ToString("yyyy-MM-ddTHH:mm:ssZ"));
        }
    }

    /// <summary>
    /// Status of the owner (draft, final, etc.)
    /// </summary>
    public string? Status
    {
        get => GetValue(DocumentPropertyName.Status);
        set => SetValue(DocumentPropertyName.Status, value ?? string.Empty);
    }
}