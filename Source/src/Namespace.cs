using System;
using System.Xml;
using System.Xml.Linq;

namespace DXPlus
{
    internal static class Name
    {
        public static XName Id = Namespace.Main + "id";
        public static XName NameId = Namespace.Main + "name";
        public static XName Paragraph = Namespace.Main + "p";
        public static XName Indent = Namespace.Main + "ind";
        public static XName ParagraphProperties = Namespace.Main + "pPr";
        public static XName ParagraphStyle = Namespace.Main + "pStyle";
        public static XName MainVal = Namespace.Main + "val";
        public static XName Run = Namespace.Main + "r";
        public static XName Text = Namespace.Main + "t";
        public static XName ParagraphAlignment = Namespace.Main + "jc";
        public static XName Bold = Namespace.Main + "b";
        public static XName Italic = Namespace.Main + "i";
        public static XName Underline = Namespace.Main + "u";
        public static XName Emphasis = Namespace.Main + "em";
        public static XName Size = Namespace.Main + "sz";
        public static XName Color = Namespace.Main + "color";
        public static XName Language = Namespace.Main + "lang";
        public static XName RunFonts = Namespace.Main + "rFonts";
        public static XName RTL = Namespace.Main + "bidi";
        public static XName Vanish = Namespace.Main + "vanish";
        public static XName Highlight = Namespace.Main + "highlight";
        public static XName Kerning = Namespace.Main + "kern";
        public static XName SimpleField = Namespace.Main + "fldSimple";
        public static XName Left = Namespace.Main + "left";
        public static XName Right = Namespace.Main + "right";
        public static XName Top = Namespace.Main + "top";
        public static XName Bottom = Namespace.Main + "bottom";
        public static XName FirstLine = Namespace.Main + "firstLine";
        public static XName Hanging = Namespace.Main + "hanging";
        public static XName KeepNext = Namespace.Main + "keepNext";
        public static XName KeepLines = Namespace.Main + "keepLines";
        public static XName Position = Namespace.Main + "position";
        public static XName RunProperties = Namespace.Main + "rPr";
        public static XName Spacing = Namespace.Main + "spacing";
        public static XName VerticalAlign = Namespace.Main + "vertAlign";
        public static XName BookmarkStart = Namespace.Main + "bookmarkStart";
        public static XName BookmarkEnd = Namespace.Main + "bookmarkEnd";
        public static XName MathParagraph = Namespace.Math + "oMathPara";
        public static XName OfficeMath = Namespace.Math + "oMath";
        public static XName NoProof => Namespace.Main + "noProof";
    }

    internal static class Relations
    {
        public static Relationship Header = new Relationship($"{Namespace.RelatedDoc.NamespaceName}/header", "/word/header{0}.xml", DocxContentType.Base + "header+xml");
        public static Relationship Footer = new Relationship($"{Namespace.RelatedDoc.NamespaceName}/footer", "/word/footer{0}.xml", DocxContentType.Base + "footer+xml");
        public static Relationship Endnotes = new Relationship($"{Namespace.RelatedDoc.NamespaceName}/endnotes", "/word/endnotes.xml", DocxContentType.Base + "endnotes+xml");
        public static Relationship Footnotes = new Relationship($"{Namespace.RelatedDoc.NamespaceName}/footnotes", "/word/footnotes.xml", DocxContentType.Base + "footnotes+xml");
        public static Relationship Styles = new Relationship($"{Namespace.RelatedDoc.NamespaceName}/styles", "/word/styles.xml", DocxContentType.Base + "styles+xml");
        public static Relationship StylesWithEffects = new Relationship("http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects", "/word/stylesWithEffects.xml", "application/vnd.ms-word.stylesWithEffects+xml");
        public static Relationship FontTable = new Relationship($"{Namespace.RelatedDoc.NamespaceName}/fontTable", "/word/fontTable.xml", DocxContentType.Base + "fontTable+xml");
        public static Relationship Numbering = new Relationship($"{Namespace.RelatedDoc.NamespaceName}/numbering", "/word/numbering.xml", DocxContentType.Base + "numbering+xml");
        public static Relationship Settings = new Relationship($"{Namespace.RelatedDoc.NamespaceName}/settings", "/word/settings.xml", DocxContentType.Base + "settings+xml");
    }

    internal class Relationship
    {
        public string Path { get; }
        public string RelType { get; }
        public string ContentType { get; }
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

    internal static class Namespace
    {
        internal static XNamespace Main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        internal static XNamespace RelatedPackage = "http://schemas.openxmlformats.org/package/2006/relationships";
        internal static XNamespace Math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        internal static XNamespace CustomPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        internal static XNamespace CustomVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        internal static XNamespace WordProcessingDrawing = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        internal static XNamespace DrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";
        internal static XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        internal static XNamespace VML = "urn:schemas-microsoft-com:vml";
        internal static XNamespace RelatedDoc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        static XmlNamespaceManager _manager;
        internal static XmlNamespaceManager NamespaceManager()
        {
            if (_manager == null) {
                _manager = new XmlNamespaceManager(new NameTable());
                _manager.AddNamespace("w", Main.NamespaceName); // default
                _manager.AddNamespace("r", RelatedDoc.NamespaceName);
                _manager.AddNamespace("v", VML.NamespaceName);
                _manager.AddNamespace("a", DrawingMain.NamespaceName);
                _manager.AddNamespace("c", Chart.NamespaceName);
                _manager.AddNamespace("wp", WordProcessingDrawing.NamespaceName);
            }
            return _manager;
        }
    }

    internal static class DocxContentType
    {
        internal const string Base = "application/vnd.openxmlformats-officedocument.wordprocessingml.";
        internal const string Document = Base + "document.main+xml";
        internal const string Template = Base + "template.main+xml";
        internal const string Relationships = "application/vnd.openxmlformats-package.relationships+xml";
    }

    internal static class DocxSections
    {
        internal static Uri DocPropsCoreUri = new Uri("/docProps/core.xml", UriKind.Relative);
        internal static Uri DocPropsCustom = new Uri("/docProps/custom.xml", UriKind.Relative);
    }
}