using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Designs\Styles that can be applied to a table.
/// </summary>
public enum TableDesign
{
    /// <summary>
    /// Custom table style
    /// </summary>
    [XmlAttribute("Custom")] Custom,

    /// <summary>
    /// Normal
    /// </summary>
    [XmlAttribute("TableNormal")] Normal,

    /// <summary>
    /// Grid
    /// </summary>
    [XmlAttribute("TableGrid")] Grid,

    /// <summary>
    /// Light grid
    /// </summary>
    [XmlAttribute("TableGridLight")] GridLight,

    /// <summary>
    /// Light shading
    /// </summary>
    [XmlAttribute("LightShading")] LightShading,

    /// <summary>
    /// Light shading - Accent1
    /// </summary>
    [XmlAttribute("LightShading-Accent1")] LightShadingAccent1,

    /// <summary>
    /// Light shading - Accent2
    /// </summary>
    [XmlAttribute("LightShading-Accent2")] LightShadingAccent2,

    /// <summary>
    /// Light shading - Accent3
    /// </summary>
    [XmlAttribute("LightShading-Accent3")] LightShadingAccent3,

    /// <summary>
    /// Light shading - Accent4
    /// </summary>
    [XmlAttribute("LightShading-Accent4")] LightShadingAccent4,

    /// <summary>
    /// Light shading - Accent5
    /// </summary>
    [XmlAttribute("LightShading-Accent5")] LightShadingAccent5,

    /// <summary>
    /// Light shading - Accent6
    /// </summary>
    [XmlAttribute("LightShading-Accent6")] LightShadingAccent6,

    /// <summary>
    /// Light list
    /// </summary>
    [XmlAttribute("LightList")] LightList,

    /// <summary>
    /// Light list - Accent1
    /// </summary>
    [XmlAttribute("LightList-Accent1")] LightListAccent1,

    /// <summary>
    /// Light list - Accent2
    /// </summary>
    [XmlAttribute("LightList-Accent2")] LightListAccent2,

    /// <summary>
    /// Light list - Accent3
    /// </summary>
    [XmlAttribute("LightList-Accent3")] LightListAccent3,

    /// <summary>
    /// Light list - Accent4
    /// </summary>
    [XmlAttribute("LightList-Accent4")] LightListAccent4,

    /// <summary>
    /// Light list - Accent5
    /// </summary>
    [XmlAttribute("LightList-Accent5")] LightListAccent5,

    /// <summary>
    /// Light list - Accent6
    /// </summary>
    [XmlAttribute("LightList-Accent6")] LightListAccent6,

    /// <summary>
    /// Light grid
    /// </summary>
    [XmlAttribute("LightGrid")] LightGrid,

    /// <summary>
    /// Light grid - Accent1
    /// </summary>
    [XmlAttribute("LightGrid-Accent1")] LightGridAccent1,

    /// <summary>
    /// Light grid - Accent2
    /// </summary>
    [XmlAttribute("LightGrid-Accent2")] LightGridAccent2,

    /// <summary>
    /// Light grid - Accent3
    /// </summary>
    [XmlAttribute("LightGrid-Accent3")] LightGridAccent3,

    /// <summary>
    /// Light grid - Accent4
    /// </summary>
    [XmlAttribute("LightGrid-Accent4")] LightGridAccent4,

    /// <summary>
    /// Light grid - Accent5
    /// </summary>
    [XmlAttribute("LightGrid-Accent5")] LightGridAccent5,

    /// <summary>
    /// Light grid - Accent6
    /// </summary>
    [XmlAttribute("LightGrid-Accent6")] LightGridAccent6,

    /// <summary>
    /// Medium shading
    /// </summary>
    [XmlAttribute("MediumShading1")] MediumShading1,

    /// <summary>
    /// Medium shading - Accent1
    /// </summary>
    [XmlAttribute("MediumShading1-Accent1")] MediumShading1Accent1,

    /// <summary>
    /// Medium shading - Accent2
    /// </summary>
    [XmlAttribute("MediumShading1-Accent2")] MediumShading1Accent2,

    /// <summary>
    /// Medium shading - Accent3
    /// </summary>
    [XmlAttribute("MediumShading1-Accent3")] MediumShading1Accent3,

    /// <summary>
    /// Medium shading - Accent4
    /// </summary>
    [XmlAttribute("MediumShading1-Accent4")] MediumShading1Accent4,

    /// <summary>
    /// Medium shading - Accent5
    /// </summary>
    [XmlAttribute("MediumShading1-Accent5")] MediumShading1Accent5,

    /// <summary>
    /// Medium shading - Accent6
    /// </summary>
    [XmlAttribute("MediumShading1-Accent6")] MediumShading1Accent6,

    /// <summary>
    /// Medium shading 2
    /// </summary>
    [XmlAttribute("MediumShading2")] MediumShading2,

    /// <summary>
    /// Medium shading 2 - Accent1
    /// </summary>
    [XmlAttribute("MediumShading2-Accent1")] MediumShading2Accent1,

    /// <summary>
    /// Medium shading 2 - Accent2
    /// </summary>
    [XmlAttribute("MediumShading2-Accent2")] MediumShading2Accent2,

    /// <summary>
    /// Medium shading 2 - Accent3
    /// </summary>
    [XmlAttribute("MediumShading2-Accent3")] MediumShading2Accent3,

    /// <summary>
    /// Medium shading 2 - Accent4
    /// </summary>
    [XmlAttribute("MediumShading2-Accent4")] MediumShading2Accent4,

    /// <summary>
    /// Medium shading 2 - Accent5
    /// </summary>
    [XmlAttribute("MediumShading2-Accent5")] MediumShading2Accent5,

    /// <summary>
    /// Medium shading 2 - Accent6
    /// </summary>
    [XmlAttribute("MediumShading2-Accent6")] MediumShading2Accent6,

    /// <summary>
    /// Medium list
    /// </summary>
    [XmlAttribute("MediumList1")] MediumList1,

    /// <summary>
    /// Medium list - Accent1
    /// </summary>
    [XmlAttribute("MediumList1-Accent1")] MediumList1Accent1,

    /// <summary>
    /// Medium list - Accent2
    /// </summary>
    [XmlAttribute("MediumList1-Accent2")] MediumList1Accent2,

    /// <summary>
    /// Medium list - Accent3
    /// </summary>
    [XmlAttribute("MediumList1-Accent3")] MediumList1Accent3,

    /// <summary>
    /// Medium list - Accent4
    /// </summary>
    [XmlAttribute("MediumList1-Accent4")] MediumList1Accent4,

    /// <summary>
    /// Medium list - Accent5
    /// </summary>
    [XmlAttribute("MediumList1-Accent5")] MediumList1Accent5,

    /// <summary>
    /// Medium list - Accent6
    /// </summary>
    [XmlAttribute("MediumList1-Accent6")] MediumList1Accent6,

    /// <summary>
    /// Medium list 2
    /// </summary>
    [XmlAttribute("MediumList2")] MediumList2,

    /// <summary>
    /// Medium list 2 - Accent1
    /// </summary>
    [XmlAttribute("MediumList2-Accent1")] MediumList2Accent1,

    /// <summary>
    /// Medium list 2 - Accent2
    /// </summary>
    [XmlAttribute("MediumList2-Accent2")] MediumList2Accent2,

    /// <summary>
    /// Medium list 2 - Accent3
    /// </summary>
    [XmlAttribute("MediumList2-Accent3")] MediumList2Accent3,

    /// <summary>
    /// Medium list 2 - Accent4
    /// </summary>
    [XmlAttribute("MediumList2-Accent4")] MediumList2Accent4,

    /// <summary>
    /// Medium list 2 - Accent5
    /// </summary>
    [XmlAttribute("MediumList2-Accent5")] MediumList2Accent5,

    /// <summary>
    /// Medium list 2 - Accent6
    /// </summary>
    [XmlAttribute("MediumList2-Accent6")] MediumList2Accent6,

    /// <summary>
    /// Medium grid
    /// </summary>
    [XmlAttribute("MediumGrid1")] MediumGrid1,

    /// <summary>
    /// Medium grid - Accent1
    /// </summary>
    [XmlAttribute("MediumGrid1-Accent1")] MediumGrid1Accent1,

    /// <summary>
    /// Medium grid - Accent2
    /// </summary>
    [XmlAttribute("MediumGrid1-Accent2")] MediumGrid1Accent2,

    /// <summary>
    /// Medium grid - Accent3
    /// </summary>
    [XmlAttribute("MediumGrid1-Accent3")] MediumGrid1Accent3,

    /// <summary>
    /// Medium grid - Accent4
    /// </summary>
    [XmlAttribute("MediumGrid1-Accent4")] MediumGrid1Accent4,

    /// <summary>
    /// Medium grid - Accent5
    /// </summary>
    [XmlAttribute("MediumGrid1-Accent5")] MediumGrid1Accent5,

    /// <summary>
    /// Medium grid - Accent6
    /// </summary>
    [XmlAttribute("MediumGrid1-Accent6")] MediumGrid1Accent6,

    /// <summary>
    /// Medium grid 2
    /// </summary>
    [XmlAttribute("MediumGrid2")] MediumGrid2,

    /// <summary>
    /// Medium grid 2 - Accent1
    /// </summary>
    [XmlAttribute("MediumGrid2-Accent1")] MediumGrid2Accent1,

    /// <summary>
    /// Medium grid 2 - Accent2
    /// </summary>
    [XmlAttribute("MediumGrid2-Accent2")] MediumGrid2Accent2,

    /// <summary>
    /// Medium grid 2 - Accent3
    /// </summary>
    [XmlAttribute("MediumGrid2-Accent3")] MediumGrid2Accent3,

    /// <summary>
    /// Medium grid 2 - Accent4
    /// </summary>
    [XmlAttribute("MediumGrid2-Accent4")] MediumGrid2Accent4,

    /// <summary>
    /// Medium grid 2 - Accent5
    /// </summary>
    [XmlAttribute("MediumGrid2-Accent5")] MediumGrid2Accent5,

    /// <summary>
    /// Medium grid 2 - Accent6
    /// </summary>
    [XmlAttribute("MediumGrid2-Accent6")] MediumGrid2Accent6,

    /// <summary>
    /// Medium grid 3
    /// </summary>
    [XmlAttribute("MediumGrid3")] MediumGrid3,

    /// <summary>
    /// Medium grid 3 - Accent1
    /// </summary>
    [XmlAttribute("MediumGrid3-Accent1")] MediumGrid3Accent1,

    /// <summary>
    /// Medium grid 3 - Accent2
    /// </summary>
    [XmlAttribute("MediumGrid3-Accent2")] MediumGrid3Accent2,

    /// <summary>
    /// Medium grid 3 - Accent3
    /// </summary>
    [XmlAttribute("MediumGrid3-Accent3")] MediumGrid3Accent3,

    /// <summary>
    /// Medium grid 3 - Accent4
    /// </summary>
    [XmlAttribute("MediumGrid3-Accent4")] MediumGrid3Accent4,

    /// <summary>
    /// Medium grid 3 - Accent5
    /// </summary>
    [XmlAttribute("MediumGrid3-Accent5")] MediumGrid3Accent5,

    /// <summary>
    /// Medium grid 3 - Accent6
    /// </summary>
    [XmlAttribute("MediumGrid3-Accent6")] MediumGrid3Accent6,

    /// <summary>
    /// Dark list
    /// </summary>
    [XmlAttribute("DarkList")] DarkList,

    /// <summary>
    /// Dark list - Accent1
    /// </summary>
    [XmlAttribute("DarkList-Accent1")] DarkListAccent1,

    /// <summary>
    /// Dark list - Accent2
    /// </summary>
    [XmlAttribute("DarkList-Accent2")] DarkListAccent2,

    /// <summary>
    /// Dark list - Accent3
    /// </summary>
    [XmlAttribute("DarkList-Accent3")] DarkListAccent3,

    /// <summary>
    /// Dark list - Accent4
    /// </summary>
    [XmlAttribute("DarkList-Accent4")] DarkListAccent4,

    /// <summary>
    /// Dark list - Accent5
    /// </summary>
    [XmlAttribute("DarkList-Accent5")] DarkListAccent5,

    /// <summary>
    /// Dark list - Accent6
    /// </summary>
    [XmlAttribute("DarkList-Accent6")] DarkListAccent6,

    /// <summary>
    /// Colorful shading
    /// </summary>
    [XmlAttribute("ColorfulShading")] ColorfulShading,

    /// <summary>
    /// Colorful shading - Accent1
    /// </summary>
    [XmlAttribute("ColorfulShading-Accent1")] ColorfulShadingAccent1,

    /// <summary>
    /// Colorful shading - Accent2
    /// </summary>
    [XmlAttribute("ColorfulShading-Accent2")] ColorfulShadingAccent2,

    /// <summary>
    /// Colorful shading - Accent3
    /// </summary>
    [XmlAttribute("ColorfulShading-Accent3")] ColorfulShadingAccent3,

    /// <summary>
    /// Colorful shading - Accent4
    /// </summary>
    [XmlAttribute("ColorfulShading-Accent4")] ColorfulShadingAccent4,

    /// <summary>
    /// Colorful shading - Accent5
    /// </summary>
    [XmlAttribute("ColorfulShading-Accent5")] ColorfulShadingAccent5,

    /// <summary>
    /// Colorful shading - Accent6
    /// </summary>
    [XmlAttribute("ColorfulShading-Accent6")] ColorfulShadingAccent6,

    /// <summary>
    /// Colorful list
    /// </summary>
    [XmlAttribute("ColorfulList")] ColorfulList,

    /// <summary>
    /// Colorful list - Accent1
    /// </summary>
    [XmlAttribute("ColorfulList-Accent1")] ColorfulListAccent1,

    /// <summary>
    /// Colorful list - Accent2
    /// </summary>
    [XmlAttribute("ColorfulList-Accent2")] ColorfulListAccent2,

    /// <summary>
    /// Colorful list - Accent3
    /// </summary>
    [XmlAttribute("ColorfulList-Accent3")] ColorfulListAccent3,

    /// <summary>
    /// Colorful list - Accent4
    /// </summary>
    [XmlAttribute("ColorfulList-Accent4")] ColorfulListAccent4,

    /// <summary>
    /// Colorful list - Accent5
    /// </summary>
    [XmlAttribute("ColorfulList-Accent5")] ColorfulListAccent5,

    /// <summary>
    /// Colorful list - Accent6
    /// </summary>
    [XmlAttribute("ColorfulList-Accent6")] ColorfulListAccent6,

    /// <summary>
    /// Colorful grid
    /// </summary>
    [XmlAttribute("ColorfulGrid")] ColorfulGrid,

    /// <summary>
    /// Colorful grid - Accent1
    /// </summary>
    [XmlAttribute("ColorfulGrid-Accent1")] ColorfulGridAccent1,

    /// <summary>
    /// Colorful grid - Accent2
    /// </summary>
    [XmlAttribute("ColorfulGrid-Accent2")] ColorfulGridAccent2,

    /// <summary>
    /// Colorful grid - Accent3
    /// </summary>
    [XmlAttribute("ColorfulGrid-Accent3")] ColorfulGridAccent3,

    /// <summary>
    /// Colorful grid - Accent4
    /// </summary>
    [XmlAttribute("ColorfulGrid-Accent4")] ColorfulGridAccent4,

    /// <summary>
    /// Colorful grid - Accent5
    /// </summary>
    [XmlAttribute("ColorfulGrid-Accent5")] ColorfulGridAccent5,

    /// <summary>
    /// Colorful grid - Accent6
    /// </summary>
    [XmlAttribute("ColorfulGrid-Accent6")] ColorfulGridAccent6,

    /// <summary>
    /// Plain table
    /// </summary>
    [XmlAttribute("None")] None
};