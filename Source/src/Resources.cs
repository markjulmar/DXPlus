using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Xml.Linq;

namespace DXPlus
{
    internal static class Resources
    {
        public static string TocXmlBase => GetXMLResource("DXPlus.Resources.TocXmlBase.xml");
        public static string TocHeadingStyleBase => GetXMLResource("DXPlus.Resources.TocHeadingStyleBase.xml");
        public static string TocElementStyleBase => GetXMLResource("DXPlus.Resources.TocElementStyleBase.xml");
        public static string TocHyperLinkStyleBase => GetXMLResource("DXPlus.Resources.TocHyperLinkStyleBase.xml");

        public static XDocument NumberingXml => DecompressXMLResource("DXPlus.Resources.numbering.xml.gz");
        public static XDocument DefaultStylesXml => DecompressXMLResource("DXPlus.Resources.default_styles.xml.gz");
        public static XDocument DefaultBulletNumberingXml => DecompressXMLResource("DXPlus.Resources.numbering.default_bullet_abstract.xml.gz");
        public static XDocument DefaultDecimalNumberingXml => DecompressXMLResource("DXPlus.Resources.numbering.default_decimal_abstract.xml.gz");
        public static XDocument DefaultTableStyles => DecompressXMLResource("DXPlus.Resources.styles.xml.gz");

        private static string GetXMLResource(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream stream = assembly.GetManifestResourceStream(resourceName);
            using StreamReader sr = new StreamReader(stream);
            return sr.ReadToEnd();
        }

        private static XDocument DecompressXMLResource(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream stream = assembly.GetManifestResourceStream(resourceName);

            using GZipStream zip = new GZipStream(stream, CompressionMode.Decompress);
            using TextReader sr = new StreamReader(zip);
            return XDocument.Load(sr);
        }
    }
}