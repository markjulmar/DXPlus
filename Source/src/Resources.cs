using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace DXPlus
{
    internal static class Resources
    {
        public static XElement TocXmlBase(string headerStyle, string title, int? rightTabPos, string switches)
                => GetElement("DXPlus.Resources.TocXmlBase.xml", new { headerStyle, title, rightTabPos, switches });

        public static XElement TocHeadingStyleBase(string headerStyle, string name)
                => GetElement("DXPlus.Resources.TocHeadingStyleBase.xml", new { name });

        public static XElement TocElementStyleBase(string headerStyle, string name)
                => GetElement("DXPlus.Resources.TocElementStyleBase.xml", new { headerStyle, name });

        public static XElement TocHyperLinkStyleBase(string headerStyle, string name)
                => GetElement("DXPlus.Resources.TocHyperLinkStyleBase.xml");

        public static XDocument NumberingXml() => GetDocument("DXPlus.Resources.numbering.xml.gz");
        public static XDocument DefaultStylesXml() => GetDocument("DXPlus.Resources.default_styles.xml.gz");
        public static XDocument DefaultBulletNumberingXml() => GetDocument("DXPlus.Resources.numbering.default_bullet_abstract.xml.gz");
        public static XDocument DefaultDecimalNumberingXml() => GetDocument("DXPlus.Resources.numbering.default_decimal_abstract.xml.gz");
        public static XDocument DefaultTableStyles() => GetDocument("DXPlus.Resources.styles.xml.gz");
        public static XDocument SettingsXml(string rsid) => GetDocument("DXPlus.Resources.settings.xml", new { rsid });

        static XDocument GetDocument(string resourceName, object tokens = null) => XDocument.Parse(GetText(resourceName, tokens));
        static XElement GetElement(string resourceName, object tokens = null) => XElement.Parse(GetText(resourceName, tokens));

        static string GetText(string resourceName, object tokens)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream = assembly.GetManifestResourceStream(resourceName);
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
                using StreamReader sr = new StreamReader(stream);
                text = sr.ReadToEnd();
            }

            return ReplaceTokens(text, values);
        }

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

        private static string ReplaceTokens(string text, IDictionary<string, object> tokens)
        {
            if (tokens.Count == 0)
                return text;

            var sb = new StringBuilder(text);
            foreach (var item in tokens)
            {
                string key = "{" + item.Key + "}";
                string value = item.Value?.ToString();

                sb.Replace(key, value);
            }

            return sb.ToString();
        }
    }
}