// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

namespace Sbroenne.ExcelMcp.ComInterop.Formatting;

/// <summary>
/// Local Power Query M code formatter that works offline.
/// Provides basic formatting with indentation and line breaks without external API calls.
/// </summary>
/// <remarks>
/// <para><b>Design Principles:</b></para>
/// <list type="bullet">
/// <item>Works completely offline - no network calls</item>
/// <item>Simple rule-based formatting for let/in structures</item>
/// <item>Graceful fallback - returns original on any failure</item>
/// </list>
/// <para><b>Formatting Rules:</b></para>
/// <list type="number">
/// <item>Break after 'let' keyword</item>
/// <item>Each variable assignment on new line</item>
/// <item>Break before 'in' keyword</item>
/// <item>Indent nested expressions</item>
/// <item>Break after commas in function calls</item>
/// </list>
/// </remarks>
public static class LocalMCodeFormatter
{
    private static readonly string[] s_whitespaceChars = { " ", "\t" };
    private const string Indent = "    "; // 4 spaces

    // Pre-allocated indent cache (0-20 levels) for performance
    private static readonly string[] s_indents = new string[21];
    static LocalMCodeFormatter()
    {
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < s_indents.Length; i++)
        {
            s_indents[i] = sb.ToString();
            sb.Append(Indent);
        }
    }

    private static string GetIndent(int level) =>
        level < s_indents.Length ? s_indents[level] : new string(' ', level * 4);

    /// <summary>
    /// Formats Power Query M code using simple rule-based formatting.
    /// </summary>
    /// <param name="mCode">The M code to format</param>
    /// <returns>Formatted M code, or original code if formatting fails</returns>
    public static Task<string> FormatAsync(string mCode)
    {
        if (string.IsNullOrWhiteSpace(mCode))
            return Task.FromResult(mCode);

        try
        {
            var formatted = FormatInternal(mCode.Trim());
            return Task.FromResult(formatted);
        }
        catch (Exception)
        {
            // Never throw - return original on failure
            return Task.FromResult(mCode);
        }
    }

    private static string FormatInternal(string mCode)
    {
        // Step 1: Normalize whitespace
        mCode = NormalizeWhitespace(mCode);

        // Step 2: Add line breaks and indentation
        return AddLineBreaksAndIndentation(mCode);
    }

    private static string NormalizeWhitespace(string mCode)
    {
        // Replace multiple spaces/tabs with single space
        var parts = mCode.Split(s_whitespaceChars, StringSplitOptions.RemoveEmptyEntries);
        return string.Join(" ", parts);
    }

    private static string AddLineBreaksAndIndentation(string mCode)
    {
        var result = new System.Text.StringBuilder();
        var indentLevel = 0;
        var inLetSection = false;

        for (int i = 0; i < mCode.Length; i++)
        {
            var ch = mCode[i];
            var remaining = mCode.Substring(i);

            // Check for 'let' keyword
            if (remaining.StartsWith("let ", StringComparison.OrdinalIgnoreCase) ||
                remaining.Equals("let", StringComparison.OrdinalIgnoreCase))
            {
                result.Append("let");
                result.AppendLine();
                indentLevel = 1;
                inLetSection = true;
                i += 2; // Skip 'let'
                continue;
            }

            // Check for 'in' keyword
            if ((remaining.StartsWith("in ", StringComparison.OrdinalIgnoreCase) ||
                 remaining.Equals("in", StringComparison.OrdinalIgnoreCase)) && inLetSection)
            {
                indentLevel = Math.Max(0, indentLevel - 1);
                result.AppendLine();
                result.Append(GetIndent(indentLevel));
                result.Append("in");
                inLetSection = false;
                i += 1; // Skip 'in'
                continue;
            }

            switch (ch)
            {
                case '(':
                case '[':
                case '{':
                    result.Append(ch);
                    indentLevel++;
                    result.AppendLine();
                    result.Append(GetIndent(indentLevel));
                    break;

                case ')':
                case ']':
                case '}':
                    indentLevel = Math.Max(0, indentLevel - 1);
                    result.AppendLine();
                    result.Append(GetIndent(indentLevel));
                    result.Append(ch);
                    break;

                case ',':
                    result.Append(',');
                    result.AppendLine();
                    result.Append(GetIndent(indentLevel));
                    break;

                case '=' when inLetSection:
                    result.Append(" = ");
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
