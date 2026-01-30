using System.Text;

namespace PandocGUI.Utilities;

public static class CommandLineTokenizer
{
    public static IReadOnlyList<string> Split(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return Array.Empty<string>();
        }

        var args = new List<string>();
        var buffer = new StringBuilder();
        var inQuotes = false;
        var quoteChar = '"';

        foreach (var ch in input)
        {
            if (ch == '"' || ch == (char)39)
            {
                if (inQuotes && ch == quoteChar)
                {
                    inQuotes = false;
                }
                else if (!inQuotes)
                {
                    inQuotes = true;
                    quoteChar = ch;
                }
                else
                {
                    buffer.Append(ch);
                }

                continue;
            }
            if (char.IsWhiteSpace(ch) && !inQuotes)
            {
                if (buffer.Length > 0)
                {
                    args.Add(buffer.ToString());
                    buffer.Clear();
                }

                continue;
            }

            buffer.Append(ch);
        }

        if (buffer.Length > 0)
        {
            args.Add(buffer.ToString());
        }

        return args;
    }
}
