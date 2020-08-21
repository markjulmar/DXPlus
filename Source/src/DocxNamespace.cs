using System;
using System.Xml.Linq;

namespace DXPlus
{
    internal static class DocxRelations
    {
        public static DocxRelationship Endnotes = new DocxRelationship($"{DocxNamespace.RelatedDoc.NamespaceName}/endnotes", "/word/endnotes.xml", DocxContentType.Base + "endnotes+xml");
        public static DocxRelationship Footnotes = new DocxRelationship($"{DocxNamespace.RelatedDoc.NamespaceName}/footnotes", "/word/footnotes.xml", DocxContentType.Base + "footnotes+xml");
        public static DocxRelationship Styles = new DocxRelationship($"{DocxNamespace.RelatedDoc.NamespaceName}/styles", "/word/styles.xml", DocxContentType.Base + "styles+xml");
        public static DocxRelationship StylesWithEffects = new DocxRelationship("http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects", "/word/stylesWithEffects.xml", "application/vnd.ms-word.stylesWithEffects+xml");
        public static DocxRelationship FontTable = new DocxRelationship($"{DocxNamespace.RelatedDoc.NamespaceName}/fontTable", "/word/fontTable.xml", DocxContentType.Base + "fontTable+xml");
        public static DocxRelationship Numbering = new DocxRelationship($"{DocxNamespace.RelatedDoc.NamespaceName}/numbering", "/word/numbering.xml", DocxContentType.Base + "numbering+xml");
        public static DocxRelationship Settings = new DocxRelationship($"{DocxNamespace.RelatedDoc.NamespaceName}/settings", "/word/settings.xml", DocxContentType.Base + "settings+xml");
    }

    internal class DocxRelationship
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
        public DocxRelationship(string relType, string path, string contentType)
        {
            Path = path;
            RelType = relType;
            ContentType = contentType;
            Uri = new Uri(path, UriKind.Relative);
        }
    }

    internal static class DocxNamespace
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