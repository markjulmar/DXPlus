using System;
using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus
{
    public class PageLayout : DocXElement
    {
        private const int A4_WIDTH = 11906;
        private const int A4_HEIGHT = 16838;

        internal PageLayout(DocX document, XElement xml)
            : base(document, xml)
        {
        }

        public Orientation Orientation
        {
            get
            {
                // Get the pgSz (page size) value + orient attribute
                string value = Xml.Element(DocxNamespace.Main + "pgSz")?
                    .AttributeValue(DocxNamespace.Main + "orient");

                return value?.Equals("landscape", StringComparison.CurrentCultureIgnoreCase) == true
                    ? Orientation.Landscape
                    : Orientation.Portrait; // default
            }

            set
            {
                // Check if already correct value.
                if (Orientation == value)
                {
                    return;
                }

                XElement pgSz = Xml.GetOrCreateElement(DocxNamespace.Main + "pgSz");
                pgSz.SetAttributeValue(DocxNamespace.Main + "orient", value.GetEnumName());

                if (value == Orientation.Landscape)
                {
                    pgSz.SetAttributeValue(DocxNamespace.Main + "w", A4_HEIGHT);
                    pgSz.SetAttributeValue(DocxNamespace.Main + "h", A4_WIDTH);
                }
                else // if (value == Orientation.Portrait)
                {
                    pgSz.SetAttributeValue(DocxNamespace.Main + "w", A4_WIDTH);
                    pgSz.SetAttributeValue(DocxNamespace.Main + "h", A4_HEIGHT);
                }
            }
        }
    }
}