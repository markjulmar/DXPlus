using System.Xml.Serialization;

namespace DXPlus;

/// <summary>
/// Flow chart shapes that can be used in the document
/// </summary>
public enum FlowchartShapes
{
    /// <summary>
    /// Process
    /// </summary>
    [XmlAttribute("flowChartProcess")] Process,

    /// <summary>
    /// Alternate process
    /// </summary>
    [XmlAttribute("flowChartAlternateProcess")] AlternateProcess,

    /// <summary>
    /// Decision
    /// </summary>
    [XmlAttribute("flowChartDecision")] Decision,

    /// <summary>
    /// I/O
    /// </summary>
    [XmlAttribute("flowChartInputOutput")] InputOutput,

    /// <summary>
    /// Predefined process
    /// </summary>
    [XmlAttribute("flowChartPredefinedProcess")] PredefinedProcess,

    /// <summary>
    /// Internal storage
    /// </summary>
    [XmlAttribute("flowChartInternalStorage")] InternalStorage,

    /// <summary>
    /// Document
    /// </summary>
    [XmlAttribute("flowChartDocument")] Document,

    /// <summary>
    /// Multiple documents
    /// </summary>
    [XmlAttribute("flowChartMultidocument")] Multidocument,

    /// <summary>
    /// Terminator
    /// </summary>
    [XmlAttribute("flowChartTerminator")] Terminator,

    /// <summary>
    /// Prepare step
    /// </summary>
    [XmlAttribute("flowChartPreparation")] Preparation,

    /// <summary>
    /// Manual input
    /// </summary>
    [XmlAttribute("flowChartManualInput")] ManualInput,

    /// <summary>
    /// Manual operation
    /// </summary>
    [XmlAttribute("flowChartManualOperation")] ManualOperation,

    /// <summary>
    /// Connector
    /// </summary>
    [XmlAttribute("flowChartConnector")] Connector,

    /// <summary>
    /// Off-page connector
    /// </summary>
    [XmlAttribute("flowChartOffpageConnector")] OffpageConnector,

    /// <summary>
    /// Punch card
    /// </summary>
    [XmlAttribute("flowChartPunchedCard")] PunchedCard,

    /// <summary>
    /// Punch tape
    /// </summary>
    [XmlAttribute("flowChartPunchedTape")] PunchedTape,

    /// <summary>
    /// Summation
    /// </summary>
    [XmlAttribute("flowChartSummingJunction")] SummingJunction,

    /// <summary>
    /// Or
    /// </summary>
    [XmlAttribute("flowChartOr")] Or,

    /// <summary>
    /// Collation
    /// </summary>
    [XmlAttribute("flowChartCollate")] Collate,

    /// <summary>
    /// Sort
    /// </summary>
    [XmlAttribute("flowChartSort")] Sort,

    /// <summary>
    /// Extract
    /// </summary>
    [XmlAttribute("flowChartExtract")] Extract,

    /// <summary>
    /// Merge
    /// </summary>
    [XmlAttribute("flowChartMerge")] Merge,

    /// <summary>
    /// Online storage
    /// </summary>
    [XmlAttribute("flowChartOnlineStorage")] OnlineStorage,

    /// <summary>
    /// Delay
    /// </summary>
    [XmlAttribute("flowChartDelay")] Delay,

    /// <summary>
    /// Magnetic tape
    /// </summary>
    [XmlAttribute("flowChartMagneticTape")] MagneticTape,

    /// <summary>
    /// Magnetic disk
    /// </summary>
    [XmlAttribute("flowChartMagneticDisk")] MagneticDisk,
    
    /// <summary>
    /// Magnetic drum
    /// </summary>
    [XmlAttribute("flowChartMagneticDrum")] MagneticDrum,

    /// <summary>
    /// Display
    /// </summary>
    [XmlAttribute("flowChartDisplay")] Display
};