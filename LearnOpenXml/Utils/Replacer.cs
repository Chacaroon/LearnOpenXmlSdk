using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace LearnOpenXml.Utils;

class Replacer
{
    public static void ReplaceKeys(OpenXmlElement element, IReadOnlyDictionary<string, string> keyValues, bool escapeKey = true)
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
        {
            foreach (var (key, value) in keyValues)
            {
                var escapedKey = EscapeKeyIfNeeded(key, escapeKey);
                
                while (paragraph.InnerText.IndexOf(escapedKey, StringComparison.Ordinal) is var index && index > -1)
                {
                    ReplaceParagraph(paragraph, index, escapedKey, value);
                }
            }
        }
    }

    public static void ReplaceKey(OpenXmlElement element, string key, string value, bool escapeKey = true)
    {
        var escapedKey = EscapeKeyIfNeeded(key, escapeKey);
        foreach (var paragraph in element.Descendants<Paragraph>())
        {
            while (paragraph.InnerText.IndexOf(escapedKey, StringComparison.Ordinal) is var index && index > -1)
            {
                ReplaceParagraph(paragraph, index, escapedKey, value);
            }
        }
    }

    // Text in a paragraph can be split by pieces of type Text
    // Trying to join pieces into one item and replace the key in it
    private static void ReplaceParagraph(OpenXmlElement paragraph, int index, string key, string value)
    {
        var count = 0;
        var builder = new StringBuilder();
        var textNodes = paragraph.Descendants<Text>().ToArray();

        if (textNodes.Length == 1)
        {
            textNodes[0].Text = textNodes[0].Text.Replace(key, value);
            return;
        }
        
        Text? textElement = null;
        foreach (var item in textNodes)
        {
            if (item.Text.Length == 0)
            {
                continue;
            }
            
            // If count is between start and end indexes of the key, it means that the key is split
            // Saving all parts of the key into a StringBuilder and removing redundant elements
            
            // For example:
            // <text>useless text just ignore</text>
            // <text>some text {key to rep</text>
            // <text>lace} more text</text>
            // 
            // after replacement
            // 
            // <text>useless text just ignore</text>
            // <text>some text REPLACED TEXT more text</text>
            if (index <= count && count < index + key.Length)
            {
                builder.Append(item.Text);

                textElement?.Remove();
                textElement = item;
            }
            
            count += item.Text.Length;
        }
        
        textElement!.Text = builder.Replace(key, value).ToString();
    }

    private static string EscapeKeyIfNeeded(string key, bool escape) => escape ? $"{{{key}}}" : key;
}