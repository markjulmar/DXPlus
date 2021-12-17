using System;

namespace DXPlus
{
    /// <summary>
    /// This is the value used on table conditional formatting (tblLook:val).
    /// </summary>
    [Flags]
    public enum TableConditionalFormatting
    {
        /// <summary>
        /// None
        /// </summary>
        None = 0,
        
        /// <summary>
        /// Apply first row conditional formatting
        /// </summary>
        FirstRow = 0x20,

        /// <summary>
        /// Apply last row conditional formatting 
        /// </summary>
        LastRow = 0x40,
        
        /// <summary>
        /// Apply first column conditional formatting
        /// </summary>
        FirstColumn = 0x80,

        /// <summary>
        /// Apply last column conditional formatting
        /// </summary>
        LastColumn = 0x100,

        /// <summary>
        /// Do not apply row banding conditional formatting
        /// </summary>
        NoRowBand = 0x200,

        /// <summary>
        /// Do not apply column banding conditional formatting
        /// </summary>
        NoColumnBand = 0x400
    }
}