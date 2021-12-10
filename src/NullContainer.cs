using System.Xml.Linq;

namespace DXPlus
{
    /// <summary>
    /// Temporary container when building up object graphs.
    /// </summary>
    internal class NullContainer : BlockContainer
    {
        public NullContainer() : base(null, new XElement(Name.Temp))
        {
        }
    }
}