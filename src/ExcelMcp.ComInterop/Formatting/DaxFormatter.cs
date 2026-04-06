// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Formats DAX (Data Analysis Expressions) code using a local offline formatter.
/// Provides automatic pretty-printing with proper indentation and line breaks.
/// </summary>
/// <remarks>
/// <para><b>Design Principles:</b></para>
/// <list type="bullet">
/// <item>Never throws exceptions - returns original DAX on any failure</item>
/// <item>Works completely offline - no external API calls</item>
/// <item>Simple rule-based formatting</item>
/// <item>Graceful fallback ensures operations never fail due to formatting</item>
/// </list>
/// <para><b>Usage:</b></para>
/// <code>
/// string formatted = await DaxFormatter.FormatAsync("CALCULATE(SUM(Sales[Amount]),FILTER(ALL(Calendar),Calendar[Year]=2024))");
/// // Returns formatted DAX with indentation, or original if formatting fails
/// </code>
/// <para><b>Performance:</b></para>
/// <list type="bullet">
/// <item>Zero network latency - completely local</item>
/// <item>Fast rule-based formatting</item>
/// <item>Graceful fallback ensures operations never break due to formatting</item>
/// </list>
/// </remarks>
public static class DaxFormatter
{
    /// <summary>
    /// Formats DAX code using a local offline formatter.
    /// </summary>
    /// <param name="daxCode">The DAX code to format</param>
    /// <param name="cancellationToken">Cancellation token (unused for local formatting)</param>
    /// <returns>Formatted DAX code, or original code if formatting fails</returns>
    /// <remarks>
    /// This method NEVER throws exceptions. If formatting fails for any reason,
    /// it returns the original code unchanged.
    /// </remarks>
    public static async Task<string> FormatAsync(string daxCode, CancellationToken cancellationToken = default)
    {
        _ = cancellationToken; // Suppress unused parameter warning
        return await LocalDaxFormatter.FormatAsync(daxCode);
    }
}


