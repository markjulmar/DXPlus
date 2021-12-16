using System;
using System.Xml;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Standard names of elements in a Word OpenXML document. 
    /// </summary>
    internal static class Name
    {
        public static readonly XName Bold = Namespace.Main + "b";
        public static readonly XName BookmarkEnd = Namespace.Main + "bookmarkEnd";
        public static readonly XName BookmarkStart = Namespace.Main + "bookmarkStart";
        public static readonly XName Bottom = Namespace.Main + "bottom";
        public static readonly XName Color = Namespace.Main + "color";
        public static readonly XName ComplexField = Namespace.Main + "fldChar";
        public static readonly XName DrawingProperties = Namespace.WordProcessingDrawing + "docPr";
        public static readonly XName Emphasis = Namespace.Main + "em";
        public static readonly XName FirstLine = Namespace.Main + "firstLine";
        public static readonly XName Hanging = Namespace.Main + "hanging";
        public static readonly XName Highlight = Namespace.Main + "highlight";
        public static readonly XName Id = Namespace.Main + "id";
        public static readonly XName Indent = Namespace.Main + "ind";
        public static readonly XName Italic = Namespace.Main + "i";
        public static readonly XName KeepLines = Namespace.Main + "keepLines";
        public static readonly XName KeepNext = Namespace.Main + "keepNext";
        public static readonly XName Kerning = Namespace.Main + "kern";
        public static readonly XName Language = Namespace.Main + "lang";
        public static readonly XName Left = Namespace.Main + "left";
        public static readonly XName MainVal = Namespace.Main + "val";
        public static readonly XName MathParagraph = Namespace.Math + "oMathPara";
        public static readonly XName NameId = Namespace.Main + "name";
        public static readonly XName NoProof = Namespace.Main + "noProof";
        public static readonly XName OfficeMath = Namespace.Math + "oMath";
        public static readonly XName Paragraph = Namespace.Main + "p";
        public static readonly XName ParagraphAlignment = Namespace.Main + "jc";
        public static readonly XName ParagraphId = Namespace.W2010 + "paraId";
        public static readonly XName ParagraphProperties = Namespace.Main + "pPr";
        public static readonly XName ParagraphStyle = Namespace.Main + "pStyle";
        public static readonly XName Position = Namespace.Main + "position";
        public static readonly XName RightToLeft = Namespace.Main + "bidi";
        public static readonly XName Right = Namespace.Main + "right";
        public static readonly XName Run = Namespace.Main + "r";
        public static readonly XName RunFonts = Namespace.Main + "rFonts";
        public static readonly XName RunProperties = Namespace.Main + "rPr";
        public static readonly XName SectionProperties = Namespace.Main + "sectPr";
        public static readonly XName Shadow = Namespace.Main + "shadow";
        public static readonly XName SimpleField = Namespace.Main + "fldSimple";
        public static readonly XName Size = Namespace.Main + "sz";
        public static readonly XName Spacing = Namespace.Main + "spacing";
        public static readonly XName Table = Namespace.Main + "tbl";
        public static readonly XName TableIndent = Namespace.Main + "tblInd";
        public static readonly XName Text = Namespace.Main + "t";
        public static readonly XName Top = Namespace.Main + "top";
        public static readonly XName Underline = Namespace.Main + "u";
        public static readonly XName Vanish = Namespace.Main + "vanish";
        public static readonly XName VerticalAlign = Namespace.Main + "vertAlign";
    }

    /// <summary>
    /// Define the comment relationships in OpenXML.
    /// </summary>
    internal static class Relations
    {
        public static readonly Relationship Header = new($"{Namespace.RelatedDoc.NamespaceName}/header", "/word/header{0}.xml", DocxContentType.Base + "header+xml");
        public static readonly Relationship Footer = new($"{Namespace.RelatedDoc.NamespaceName}/footer", "/word/footer{0}.xml", DocxContentType.Base + "footer+xml");
        public static readonly Relationship Endnotes = new($"{Namespace.RelatedDoc.NamespaceName}/endnotes", "/word/endnotes.xml", DocxContentType.Base + "endnotes+xml");
        public static readonly Relationship Footnotes = new($"{Namespace.RelatedDoc.NamespaceName}/footnotes", "/word/footnotes.xml", DocxContentType.Base + "footnotes+xml");
        public static readonly Relationship Styles = new($"{Namespace.RelatedDoc.NamespaceName}/styles", "/word/styles.xml", DocxContentType.Base + "styles+xml");
        public static readonly Relationship FontTable = new($"{Namespace.RelatedDoc.NamespaceName}/fontTable", "/word/fontTable.xml", DocxContentType.Base + "fontTable+xml");
        public static readonly Relationship Numbering = new($"{Namespace.RelatedDoc.NamespaceName}/numbering", "/word/numbering.xml", DocxContentType.Base + "numbering+xml");
        public static readonly Relationship Settings = new($"{Namespace.RelatedDoc.NamespaceName}/settings", "/word/settings.xml", DocxContentType.Base + "settings+xml");
        public static readonly Relationship CoreProperties = new($"{Namespace.RelatedPackage}/metadata/core-properties", "/docProps/core.xml", DocxContentType.CoreProperties);
        public static readonly Relationship CustomProperties = new($"{Namespace.RelatedDoc}/custom-properties", "/docProps/custom.xml", DocxContentType.CustomProperties);
        public static readonly Relationship People = new($"{Namespace.W2012Rel.NamespaceName}/people", "/word/people.xml", DocxContentType.People);
        public static readonly Relationship Comments = new($"{Namespace.RelatedDoc.NamespaceName}/comments", "/word/comments.xml", DocxContentType.Comments);
    }

    /// <summary>
    /// Represents a relationship between package parts in OpenXML
    /// </summary>
    internal class Relationship
    {
        /// <summary>
        /// Path in the package to this part
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Relationship type
        /// </summary>
        public string RelType { get; }
        
        /// <summary>
        /// Content type in the file
        /// </summary>
        public string ContentType { get; }

        /// <summary>
        /// Uri to the file (created from the path)
        /// </summary>
        public Uri Uri { get; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="relType"></param>
        /// <param name="path"></param>
        /// <param name="contentType"></param>
        public Relationship(string relType, string path, string contentType)
        {
            Path = path;
            RelType = relType;
            ContentType = contentType;
            Uri = new Uri(path, UriKind.Relative);
        }
    }

    /// <summary>
    /// Common namespaces used in the OpenXML package.
    /// </summary>
    internal static class Namespace
    {
        private static XmlNamespaceManager _manager;

        public static readonly XNamespace Main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace RelatedPackage = "http://schemas.openxmlformats.org/package/2006/relationships";
        public static readonly XNamespace Math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        public static readonly XNamespace CustomPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        public static readonly XNamespace CustomVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        public static readonly XNamespace WordProcessingDrawing = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XNamespace DrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static readonly XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        public static readonly XNamespace VML = "urn:schemas-microsoft-com:vml";
        public static readonly XNamespace RelatedDoc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace W2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
        public static readonly XNamespace ADec = "http://schemas.microsoft.com/office/drawing/2017/decorative";
        public static readonly XNamespace Picture = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        public static readonly XNamespace W2012 = "http://schemas.microsoft.com/office/word/2012/wordml";
        public static readonly XNamespace W2012Rel = "http://schemas.microsoft.com/office/2011/relationships";

        /// <summary>
        /// Returns an XmlNamespaceManager which can be used with XQuery
        /// </summary>
        public static XmlNamespaceManager NamespaceManager()
        {
            if (_manager == null) {
                _manager = new XmlNamespaceManager(new NameTable());
                _manager.AddNamespace("w", Main.NamespaceName); // default
                _manager.AddNamespace("r", RelatedDoc.NamespaceName);
                _manager.AddNamespace("v", VML.NamespaceName);
                _manager.AddNamespace("a", DrawingMain.NamespaceName);
                _manager.AddNamespace("c", Chart.NamespaceName);
                _manager.AddNamespace("w14", W2010.NamespaceName);
                _manager.AddNamespace("wp", WordProcessingDrawing.NamespaceName);
            }
            return _manager;
        }
    }

    /// <summary>
    /// Content types in the OpenXML package.
    /// </summary>
    internal static class DocxContentType
    {
        public const string Base = "application/vnd.openxmlformats-officedocument.wordprocessingml.";
        public const string Document = Base + "document.main+xml";
        public const string Template = Base + "template.main+xml";
        public const string Relationships = "application/vnd.openxmlformats-package.relationships+xml";
        public const string CoreProperties = "application/vnd.openxmlformats-package.core-properties+xml";
        public const string CustomProperties = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
        public const string People = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml";
        public const string Comments = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    }
}