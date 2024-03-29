﻿namespace DXPlus;

/// <summary>
/// Text types in a Run
/// </summary>
internal static class RunTextType
{
    /// <summary>
    /// Inserted text
    /// </summary>
    public const string InsertMarker = "ins";

    /// <summary>
    /// Deleted text
    /// </summary>
    public const string DeleteMarker = "del";

    /// <summary>
    /// Text
    /// </summary>
    public const string Text = "t";

    /// <summary>
    /// Deleted text
    /// </summary>
    public const string DeletedText = "delText";

    /// <summary>
    /// Tab
    /// </summary>
    public const string Tab = "tab";

    /// <summary>
    /// Text break
    /// </summary>
    public const string LineBreak = "br";

    /// <summary>
    /// Carriage return
    /// </summary>
    public const string CarriageReturn = "cr";

    /// <summary>
    /// Comment reference
    /// </summary>
    public const string CommentReference = "commentReference";

    /// <summary>
    /// Drawing
    /// </summary>
    public const string Drawing = "drawing";
}