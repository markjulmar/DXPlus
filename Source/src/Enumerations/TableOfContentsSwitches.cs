using System;
using System.ComponentModel;

namespace DXPlus
{
    /// <summary>
    /// Represents the switches set on a TOC.
    /// See https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_TOCTOC_topic_ID0ELZO1.html
    /// </summary>
    [Flags]
    public enum TableOfContentsSwitches
    {
        None = 0,
        [Description(@"\a")]
        A = 1,                      // Includes captioned items, but omits caption labels and numbers.
        [Description(@"\b")]
        B = 2,                      // Includes entries only from the named bookmark.
        [Description(@"\c")]
        C = 4,                      // Include figures, charts and tables.
        [Description(@"\d")]
        D = 8,                      // Separator between sequence and page#. Default is hyphen
        [Description(@"\f")]
        F = 16,                     // Include only the matching identifiers
        [Description(@"\h")]
        H = 32,                     // Uses hyperlinks
        [Description(@"\l")]
        L = 64,                     // Include specific levels
        [Description(@"\n")]
        N = 128,                    // Omit page numbers
        [Description(@"\o")]
        O = 256,                    // Use paragraphs formatted with heading styles
        [Description(@"\p")]
        P = 512,                    // Separator between entry and page#. Default is tab
        [Description(@"\s")]
        S = 1024,                   // Sequences get an added prefix to the page#.
        [Description(@"\t")]
        T = 2048,                   // Use the specific paragraph styles
        [Description(@"\u")]
        U = 4096,                   // Use the applied paragraph outline level
        [Description(@"\w")]
        W = 8192,                   // Preserve tab entries within table entries
        [Description(@"\x")]
        X = 16384,                  // Preserve newline characters within table entries
        [Description(@"\z")]
        Z = 32768                   // Hides tab leader and page numbers in Web layout view
    }
}