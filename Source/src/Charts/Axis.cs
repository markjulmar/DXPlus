using System.Xml.Linq;
using DXPlus.Helpers;

namespace DXPlus.Charts
{
    /// <summary>
    /// Axis base class
    /// </summary>
    public abstract class Axis
    {
        protected Axis()
        {
        }

        protected Axis(XElement xml)
        {
            Xml = xml;
        }

        /// <summary>
        /// ID of this Axis
        /// </summary>
        public string Id => Xml.Element(DocxNamespace.Chart + "axId").GetVal();

        /// <summary>
        /// Return true if this axis is visible
        /// </summary>
        public bool IsVisible
        {
            get => Xml.Element(DocxNamespace.Chart + "delete")?.GetVal() == "0";
            set => Xml.GetOrCreateElement(DocxNamespace.Chart + "delete").SetValue(value ? "1" : "0");
        }

        /// <summary>
        /// Axis xml element
        /// </summary>
        internal XElement Xml { get; set; }
    }
}