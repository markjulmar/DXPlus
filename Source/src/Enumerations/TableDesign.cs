﻿using System.Xml.Serialization;

namespace DXPlus
{

    /// <summary>
    /// Designs\Styles that can be applied to a table.
    /// </summary>
    public enum TableDesign
    {
        [XmlAttribute("Custom")]
        Custom,
        [XmlAttribute("TableNormal")]
        TableNormal,
        [XmlAttribute("TableGrid")]
        TableGrid,
        [XmlAttribute("LightShading")]
        LightShading,
        [XmlAttribute("LightShading-Accent1")]
        LightShadingAccent1,
        [XmlAttribute("LightShading-Accent2")]
        LightShadingAccent2,
        [XmlAttribute("LightShading-Accent3")]
        LightShadingAccent3,
        [XmlAttribute("LightShading-Accent4")]
        LightShadingAccent4,
        [XmlAttribute("LightShading-Accent5")]
        LightShadingAccent5,
        [XmlAttribute("LightShading-Accent6")]
        LightShadingAccent6,
        [XmlAttribute("LightList")]
        LightList,
        [XmlAttribute("LightList-Accent1")]
        LightListAccent1,
        [XmlAttribute("LightList-Accent2")]
        LightListAccent2,
        [XmlAttribute("LightList-Accent3")]
        LightListAccent3,
        [XmlAttribute("LightList-Accent4")]
        LightListAccent4,
        [XmlAttribute("LightList-Accent5")]
        LightListAccent5,
        [XmlAttribute("LightList-Accent6")]
        LightListAccent6,
        [XmlAttribute("LightGrid")]
        LightGrid,
        [XmlAttribute("LightGrid-Accent1")]
        LightGridAccent1,
        [XmlAttribute("LightGrid-Accent2")]
        LightGridAccent2,
        [XmlAttribute("LightGrid-Accent3")]
        LightGridAccent3,
        [XmlAttribute("LightGrid-Accent4")]
        LightGridAccent4,
        [XmlAttribute("LightGrid-Accent5")]
        LightGridAccent5,
        [XmlAttribute("LightGrid-Accent6")]
        LightGridAccent6,
        [XmlAttribute("MediumShading1")]
        MediumShading1,
        [XmlAttribute("MediumShading1-Accent1")]
        MediumShading1Accent1,
        [XmlAttribute("MediumShading1-Accent2")]
        MediumShading1Accent2,
        [XmlAttribute("MediumShading1-Accent3")]
        MediumShading1Accent3,
        [XmlAttribute("MediumShading1-Accent4")]
        MediumShading1Accent4,
        [XmlAttribute("MediumShading1-Accent5")]
        MediumShading1Accent5,
        [XmlAttribute("MediumShading1-Accent6")]
        MediumShading1Accent6,
        [XmlAttribute("MediumShading2")]
        MediumShading2,
        [XmlAttribute("MediumShading2-Accent1")]
        MediumShading2Accent1,
        [XmlAttribute("MediumShading2-Accent2")]
        MediumShading2Accent2,
        [XmlAttribute("MediumShading2-Accent3")]
        MediumShading2Accent3,
        [XmlAttribute("MediumShading2-Accent4")]
        MediumShading2Accent4,
        [XmlAttribute("MediumShading2-Accent5")]
        MediumShading2Accent5,
        [XmlAttribute("MediumShading2-Accent6")]
        MediumShading2Accent6,
        [XmlAttribute("MediumList1")]
        MediumList1,
        [XmlAttribute("MediumList1-Accent1")]
        MediumList1Accent1,
        [XmlAttribute("MediumList1-Accent2")]
        MediumList1Accent2,
        [XmlAttribute("MediumList1-Accent3")]
        MediumList1Accent3,
        [XmlAttribute("MediumList1-Accent4")]
        MediumList1Accent4,
        [XmlAttribute("MediumList1-Accent5")]
        MediumList1Accent5,
        [XmlAttribute("MediumList1-Accent6")]
        MediumList1Accent6,
        [XmlAttribute("MediumList2")]
        MediumList2,
        [XmlAttribute("MediumList2-Accent1")]
        MediumList2Accent1,
        [XmlAttribute("MediumList2-Accent2")]
        MediumList2Accent2,
        [XmlAttribute("MediumList2-Accent3")]
        MediumList2Accent3,
        [XmlAttribute("MediumList2-Accent4")]
        MediumList2Accent4,
        [XmlAttribute("MediumList2-Accent5")]
        MediumList2Accent5,
        [XmlAttribute("MediumList2-Accent6")]
        MediumList2Accent6,
        [XmlAttribute("MediumGrid1")]
        MediumGrid1,
        [XmlAttribute("MediumGrid1-Accent1")]
        MediumGrid1Accent1,
        [XmlAttribute("MediumGrid1-Accent2")]
        MediumGrid1Accent2,
        [XmlAttribute("MediumGrid1-Accent3")]
        MediumGrid1Accent3,
        [XmlAttribute("MediumGrid1-Accent4")]
        MediumGrid1Accent4,
        [XmlAttribute("MediumGrid1-Accent5")]
        MediumGrid1Accent5,
        [XmlAttribute("MediumGrid1-Accent6")]
        MediumGrid1Accent6,
        [XmlAttribute("MediumGrid2")]
        MediumGrid2,
        [XmlAttribute("MediumGrid2-Accent1")]
        MediumGrid2Accent1,
        [XmlAttribute("MediumGrid2-Accent2")]
        MediumGrid2Accent2,
        [XmlAttribute("MediumGrid2-Accent3")]
        MediumGrid2Accent3,
        [XmlAttribute("MediumGrid2-Accent4")]
        MediumGrid2Accent4,
        [XmlAttribute("MediumGrid2-Accent5")]
        MediumGrid2Accent5,
        [XmlAttribute("MediumGrid2-Accent6")]
        MediumGrid2Accent6,
        [XmlAttribute("MediumGrid3")]
        MediumGrid3,
        [XmlAttribute("MediumGrid3-Accent1")]
        MediumGrid3Accent1,
        [XmlAttribute("MediumGrid3-Accent2")]
        MediumGrid3Accent2,
        [XmlAttribute("MediumGrid3-Accent3")]
        MediumGrid3Accent3,
        [XmlAttribute("MediumGrid3-Accent4")]
        MediumGrid3Accent4,
        [XmlAttribute("MediumGrid3-Accent5")]
        MediumGrid3Accent5,
        [XmlAttribute("MediumGrid3-Accent6")]
        MediumGrid3Accent6,
        [XmlAttribute("DarkList")]
        DarkList,
        [XmlAttribute("DarkList-Accent1")]
        DarkListAccent1,
        [XmlAttribute("DarkList-Accent2")]
        DarkListAccent2,
        [XmlAttribute("DarkList-Accent3")]
        DarkListAccent3,
        [XmlAttribute("DarkList-Accent4")]
        DarkListAccent4,
        [XmlAttribute("DarkList-Accent5")]
        DarkListAccent5,
        [XmlAttribute("DarkList-Accent6")]
        DarkListAccent6,
        [XmlAttribute("ColorfulShading")]
        ColorfulShading,
        [XmlAttribute("ColorfulShading-Accent1")]
        ColorfulShadingAccent1,
        [XmlAttribute("ColorfulShading-Accent2")]
        ColorfulShadingAccent2,
        [XmlAttribute("ColorfulShading-Accent3")]
        ColorfulShadingAccent3,
        [XmlAttribute("ColorfulShading-Accent4")]
        ColorfulShadingAccent4,
        [XmlAttribute("ColorfulShading-Accent5")]
        ColorfulShadingAccent5,
        [XmlAttribute("ColorfulShading-Accent6")]
        ColorfulShadingAccent6,
        [XmlAttribute("ColorfulList")]
        ColorfulList,
        [XmlAttribute("ColorfulList-Accent1")]
        ColorfulListAccent1,
        [XmlAttribute("ColorfulList-Accent2")]
        ColorfulListAccent2,
        [XmlAttribute("ColorfulList-Accent3")]
        ColorfulListAccent3,
        [XmlAttribute("ColorfulList-Accent4")]
        ColorfulListAccent4,
        [XmlAttribute("ColorfulList-Accent5")]
        ColorfulListAccent5,
        [XmlAttribute("ColorfulList-Accent6")]
        ColorfulListAccent6,
        [XmlAttribute("ColorfulGrid")]
        ColorfulGrid,
        [XmlAttribute("ColorfulGrid-Accent1")]
        ColorfulGridAccent1,
        [XmlAttribute("ColorfulGrid-Accent2")]
        ColorfulGridAccent2,
        [XmlAttribute("ColorfulGrid-Accent3")]
        ColorfulGridAccent3,
        [XmlAttribute("ColorfulGrid-Accent4")]
        ColorfulGridAccent4,
        [XmlAttribute("ColorfulGrid-Accent5")]
        ColorfulGridAccent5,
        [XmlAttribute("ColorfulGrid-Accent6")]
        ColorfulGridAccent6,
        [XmlAttribute("None")]
        None
    };
}