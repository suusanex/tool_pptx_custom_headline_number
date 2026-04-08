using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxHeadlineNumbering;

public sealed class PrefixReplacer
{
    public bool Replace(
        A.Paragraph paragraph,
        Regex prefixRegex,
        string newPrefix,
        string separator,
        bool insertWhenPrefixMissing)
    {
        ArgumentNullException.ThrowIfNull(paragraph);
        ArgumentNullException.ThrowIfNull(prefixRegex);
        ArgumentNullException.ThrowIfNull(newPrefix);
        ArgumentNullException.ThrowIfNull(separator);

        var textNodes = paragraph.Descendants<A.Text>().ToList();
        if (textNodes.Count == 0)
        {
            if (!insertWhenPrefixMissing)
            {
                return false;
            }

            var run = new A.Run();
            run.Append(new A.Text(newPrefix + separator));
            paragraph.PrependChild(run);
            return true;
        }

        var original = string.Concat(textNodes.Select(textNode => textNode.Text));
        var match = prefixRegex.Match(original);
        var removeLength = 0;
        if (match.Success && match.Index == 0)
        {
            removeLength = match.Length;
        }
        else if (!insertWhenPrefixMissing)
        {
            return false;
        }

        RemoveLeadingCharacters(textNodes, removeLength);
        var firstTextNode = textNodes[0];
        firstTextNode.Text = newPrefix + separator + firstTextNode.Text;
        return true;
    }

    private static void RemoveLeadingCharacters(IReadOnlyList<A.Text> textNodes, int removeLength)
    {
        var remaining = removeLength;
        foreach (var textNode in textNodes)
        {
            if (remaining == 0)
            {
                break;
            }

            var sourceText = textNode.Text ?? string.Empty;
            if (sourceText.Length <= remaining)
            {
                textNode.Text = string.Empty;
                remaining -= sourceText.Length;
                continue;
            }

            textNode.Text = sourceText[remaining..];
            remaining = 0;
        }
    }
}
