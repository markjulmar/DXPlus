using System;
using System.Xml.Linq;

namespace DXPlus
{
    internal static class DocxNamespace
    {
        static internal XNamespace Main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static internal XNamespace RelatedPackage = "http://schemas.openxmlformats.org/package/2006/relationships";
        static internal XNamespace Math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        static internal XNamespace CustomPropertiesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        static internal XNamespace CustomVTypesSchema = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        static internal XNamespace WordProcessingDrawing = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        static internal XNamespace DrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";
        static internal XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        static internal XNamespace VML = "urn:schemas-microsoft-com:vml";
        static internal XNamespace Numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        static internal XNamespace RelatedDoc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    }

    internal static class DocxContentType
    {
        internal const string Styles = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
        internal const string Document = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        internal const string Template = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
        internal const string Settings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
        internal const string Relationships = "application/vnd.openxmlformats-package.relationships+xml";
    }

    internal static class DocxSections
    {
        internal static Uri StylesUri = new Uri("/word/styles.xml", UriKind.Relative);
        internal static Uri SettingsUri = new Uri("/word/settings.xml", UriKind.Relative);
        internal static Uri DocPropsCoreUri = new Uri("/docProps/core.xml", UriKind.Relative);
        internal static Uri DocPropsCustom = new Uri("/docProps/custom.xml", UriKind.Relative);
    }
}