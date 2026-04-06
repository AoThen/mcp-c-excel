// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Local DAX (Data Analysis Expressions) formatter that works offline.
/// Provides basic formatting with indentation and line breaks without external API calls.
/// </summary>
/// <remarks>
/// <para><b>Design Principles:</b></para>
/// <list type="bullet">
/// <item>Works completely offline - no network calls</item>
/// <item>Simple rule-based formatting</item>
/// <item>Graceful fallback - returns original on any failure</item>
/// </list>
/// <para><b>Formatting Rules:</b></para>
/// <list type="number">
/// <item>Normalize whitespace (single spaces)</item>
/// <item>Break lines after commas (function arguments)</item>
/// <item>Indent based on parenthesis depth</item>
/// <item>Place opening parenthesis on same line</item>
/// <item>Align closing parenthesis with opening</item>
/// </list>
/// </remarks>
public static class LocalDaxFormatter
{
    /// <summary>
    /// Formats DAX code using simple rule-based formatting.
    /// </summary>
    /// <param name="daxCode">The DAX code to format</param>
    /// <param name="cancellationToken">Cancellation token (unused for local formatting)</param>
    /// <returns>Formatted DAX code, or original code if formatting fails</returns>
    public static Task<string> FormatAsync(string daxCode, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(daxCode))
            return Task.FromResult(daxCode);

        try
        {
            var formatted = FormatInternal(daxCode.Trim());
            return Task.FromResult(formatted);
        }
        catch (Exception)
        {
            // Never throw - return original on failure
            return Task.FromResult(daxCode);
        }
    }

    private static string FormatInternal(string dax)
    {
        // Step 1: Normalize whitespace
        dax = NormalizeWhitespace(dax);

        // Step 2: Add line breaks and indentation
        return AddLineBreaksAndIndentation(dax);
    }

    private static string NormalizeWhitespace(string dax)
    {
        // Replace multiple spaces/tabs with single space
        var parts = dax.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        return string.Join(" ", parts);
    }

    private static string AddLineBreaksAndIndentation(string dax)
    {
        var result = new System.Text.StringBuilder();
        var indentLevel = 0;
        const string indent = "    "; // 4 spaces

        for (int i = 0; i < dax.Length; i++)
        {
            var ch = dax[i];

            switch (ch)
            {
                case '(':
                    result.Append('(');
                    indentLevel++;
                    // Add newline after ( if followed by non-trivial content
                    if (i + 1 < dax.Length && dax[i + 1] != ')')
                    {
                        result.AppendLine();
                        result.Append(indent, indentLevel);
                    }
                    break;

                case ')':
                    indentLevel = Math.Max(0, indentLevel - 1);
                    result.AppendLine();
                    result.Append(indent, indentLevel);
                    result.Append(')');
                    break;

                case ',':
                    result.Append(',');
                    result.AppendLine();
                    result.Append(indent, indentLevel);
                    break;

                case ' ':
                    // Skip spaces - we add our own
                    break;

                default:
                    result.Append(ch);
                    break;
            }
        }

        return result.ToString().Trim();
    }
}
