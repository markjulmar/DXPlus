using System;
using System.Xml.Linq;

namespace DXPlus
{
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
        internal static XNamespace Numbering = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numberingDoc";
        internal static XNamespace RelatedDoc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
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