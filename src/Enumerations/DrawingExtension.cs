using System;
using System.Linq;
using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// GUIDs used for drawing extensions in <w:drawing> objects.
    /// </summary>
    internal static class DrawingExtension
    {
        /// <summary>
        /// GUID for the decorative image extension
        /// </summary>
        public const string DecorativeImageId = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

        /// <summary>
        /// GUID for the SVG image extension
        /// </summary>
        public const string SvgExtensionId = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

        /// <summary>
        /// A UseLocalDpi element that specifies a flag indicating that the local
        /// BLIP compression setting overrides the document default compression setting.
        /// </summary>
        public const string UseLocalDpiExtensionId = "{28A0092B-C50C-407E-A947-70E740481C1C}";

 
        /// <summary>
        /// Method to scan an extension list for a specific extension id
        /// </summary>
        /// <param name="owner">List owner</param>
        /// <param name="id">ID to look for</param>
        /// <param name="create">True to create it</param>
        /// <returns>XElement of [a:ext]</returns>
        public static XElement Get(XElement owner, string id, bool create)
        {
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("Value cannot be null or empty.", nameof(id));
            if (owner == null)
                return null;

            var extList = owner.Element(Namespace.DrawingMain + "extLst");
            if (extList == null && create)
            {
                extList = new XElement(Namespace.DrawingMain + "extLst");
                owner.Add(extList);
            }

            var extension = extList?.Elements(Namespace.DrawingMain + "ext")
                .SingleOrDefault(e => e.AttributeValue("uri") == id);
            if (extension == null && create)
            {
                extension = new XElement(Namespace.DrawingMain + "ext",
                    new XAttribute("uri", id));
                extList.Add(extension);
            }

            return extension;
        }
    }
}