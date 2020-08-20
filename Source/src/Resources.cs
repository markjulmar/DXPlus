using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Provides access to embedded resources in the assembly.
    /// </summary>
    internal static class Resources
    {
        /// <summary>
        /// The TOC base XML element inserted when a new TOC is created.
        /// </summary>
        /// <param name="headerStyle">Header style</param>
        /// <param name="title">Title</param>
        /// <param name="rightTabPos">Right tab position for page #s</param>
        /// <param name="switches">TOC switches</param>
        /// <returns></returns>
        public static XElement TocXmlBase(string headerStyle, string title, int? rightTabPos, string switches)
                => GetElement("DXPlus.Resources.TocXmlBase.xml", new { headerStyle, title, rightTabPos, switches });

        /// <summary>
        /// TOC header style XML
        /// </summary>
        /// <param name="headerStyle">Heading style</param>
        /// <param name="name">Identifier</param>
        /// <returns></returns>
        public static XElement TocHeadingStyleBase(string headerStyle, string name)
                => GetElement("DXPlus.Resources.TocHeadingStyleBase.xml", new { name });

        /// <summary>
        /// TOC element style XML (each line)
        /// </summary>
        /// <param name="headerStyle">Header style</param>
        /// <param name="name">Identifier</param>
        /// <returns></returns>
        public static XElement TocElementStyleBase(string headerStyle, string name)
                => GetElement("DXPlus.Resources.TocElementStyleBase.xml", new { headerStyle, name });

        /// <summary>
        /// TOC link style XML
        /// </summary>
        /// <param name="headerStyle">Header style</param>
        /// <param name="name">Identifier</param>
        /// <returns></returns>
        public static XElement TocHyperLinkStyleBase(string headerStyle, string name)
                => GetElement("DXPlus.Resources.TocHyperLinkStyleBase.xml");

        /// <summary>
        /// Numbering document.
        /// </summary>
        /// <returns></returns>
        public static XDocument NumberingXml() => GetDocument("DXPlus.Resources.numbering.xml");

        /// <summary>
        /// Default styles document - used for all new documents
        /// </summary>
        /// <returns></returns>
        public static XDocument DefaultStylesXml() => GetDocument("DXPlus.Resources.styles.xml");

        /// <summary>
        /// Default bullet style for bulleted lists - added when a new bulleted list is used.
        /// </summary>
        /// <returns></returns>
        public static XDocument DefaultBulletNumberingXml() => GetDocument("DXPlus.Resources.numbering.bullets.xml");

        /// <summary>
        /// Default numbering style for numbered lists - added when a numbered list is used
        /// </summary>
        /// <returns></returns>
        public static XDocument DefaultDecimalNumberingXml() => GetDocument("DXPlus.Resources.numbering.decimal.xml");

        /// <summary>
        /// Default styles used for tables - added when a table is used
        /// </summary>
        /// <returns></returns>
        public static XDocument DefaultTableStyles() => GetDocument("DXPlus.Resources.table_styles.xml");

        /// <summary>
        /// Default settings document used for all new documents.
        /// </summary>
        /// <param name="rsid"></param>
        /// <returns></returns>
        public static XDocument SettingsXml(string rsid) => GetDocument("DXPlus.Resources.settings.xml", new { rsid });

        /// <summary>
        /// Get a Document from embedded resources
        /// </summary>
        /// <param name="resourceName">Resource name</param>
        /// <param name="tokens">Replacement tokens</param>
        private static XDocument GetDocument(string resourceName, object tokens = null) => XDocument.Parse(GetText(resourceName, tokens));

        /// <summary>
        /// Get an XML element (fragment) from embedded resources
        /// </summary>
        /// <param name="resourceName">Resource name</param>
        /// <param name="tokens">Replacement tokens</param>
        private static XElement GetElement(string resourceName, object tokens = null) => XElement.Parse(GetText(resourceName, tokens));

        /// <summary>
        /// Retrieve a string from embedded resources
        /// </summary>
        /// <param name="resourceName">Resource name</param>
        /// <param name="tokens">Replacement tokens</param>
        private static string GetText(string resourceName, object tokens)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
                throw new ArgumentException$$"Embedded resource {resourceName} missing from assembly", nameof(resourceName));

            var values = CreatePropertyDictionary(tokens);

            string text;
            if (resourceName.EndsWith(".gz"))
            {
                using var zip = new GZipStream(stream, CompressionMode.Decompress);
                using var sr = new StreamReader(zip);
                text = sr.ReadToEnd();
            }
            else
            {
                using var sr = new StreamReader(stream);
                text = sr.ReadToEnd();
            }

            return ReplaceTokens(text, values);
        }

        /// <summary>
        /// Creates a dictionary of name/value pairs to replace in a document from an object
        /// </summary>
        /// <param name="tokens">Object with public properties to replace</param>
        /// <returns></returns>
        private static IDictionary<string, object> CreatePropertyDictionary(object tokens)
        {
            var values = new Dictionary<string, object>();
            if (tokens != null)
            {
                var type = tokens.GetType();
                foreach (var propertyInfo in type.GetProperties())
                {
                    values.Add(propertyInfo.Name, propertyInfo.GetValue(tokens));
                }
            }

            return values;
        }

        /// <summary>
        /// Replaces a set of tokens in a passed string
        /// </summary>
        /// <param name="text">Source text</param>
        /// <param name="tokens">Tokens to replace</param>
        /// <returns>Modified string</returns>
        private static string ReplaceTokens(string text, IDictionary<string, object> tokens)
        {
            if (tokens.Count == 0)
                return text;

            var sb = new StringBuilder(text);
            foreach (var (name, obj) in tokens)
            {
                string key = "{" + name + "}";
                string value = obj?.ToString() ?? string.Empty;
                sb.Replace(key, value);
            }

            return sb.ToString();
        }
    }
}